// PPTX.ts is a rewrite of the original pptxjs library.

import _ from 'lodash';
import './assets/pptxjs.less';
import { getNumTypeNum } from './utils/numeric';
import './assets/d3.min.js';
import './assets/nv.d3.min.js';
import { StyleTable } from './Style';
import { Slide } from './Slide';
import { ResourceReader } from './ResourceReader';

export interface PPTXOptions {
    url?: string;
    file?: ArrayBuffer;
    processFullTheme?: boolean | string;
    mediaProcess?: boolean;
    container?: HTMLElement;
    fontSizeFactor?: number;
    slideFactor?: number;
}

export class PPTX {
    private options: PPTXOptions = {
        processFullTheme: true,
        mediaProcess: false,
        slideFactor: 96 / 914400,
        fontSizeFactor: 4 / 3.2,
    }

    private resourceReader: ResourceReader | null = null;
    private tableStyles = null;
    private basicInfo: {
        slides: string[];
        slideLayouts: string[];
        defaultTextStyle: any;
        width: number;
        height: number;
    } | null = null;
    // html result of convert
    htmlResultArray: { type: string, data: string, slide_num?: number, file_name?: string }[] = [];
    isDone = false;
    private postRenderDone = false;
    private convertPromise: Promise<any> | null = null;
    private styleTable = new StyleTable();

    private slides: Slide[] = [];

    constructor(options: PPTXOptions) {
        _.assign(this.options, options);
        this.convertPromise = this.convert();
    }

    async convert() {
        const url = this.options.url;
        let file = this.options.file;
        if (url) {
            try {
                const response = await fetch(url);
                file = await response.arrayBuffer();
            } catch (e) {
            }
        }
        if (file) {
            this.resourceReader = new ResourceReader(file);
            return await this.convertPPTX()
        }
    }

    private async getThumbnail() {
        return await this.resourceReader?.getFile("docProps/thumbnail.jpeg", "base64");
    }

    private async convertPPTX() {
        const post_ary = [];
        const [thubmnail, basicInfo, tableStyles] = await Promise.all([
            this.getThumbnail(),
            this.getBasicInfo(),
            this.readXmlFile("ppt/tableStyles.xml"),
        ])
        this.basicInfo = basicInfo;
        this.tableStyles = tableStyles;
        let numOfSlides = basicInfo["slides"].length;
        for (var i = 0; i < numOfSlides; i++) {
            let fileName = basicInfo["slides"][i];
            let filename_no_path = "";
            let filename_no_path_ary = [];
            if (fileName.indexOf("/") != -1) {
                filename_no_path_ary = fileName.split("/");
                filename_no_path = filename_no_path_ary.pop();
            } else {
                filename_no_path = fileName;
            }
            let filename_no_path_no_ext = "";
            if (filename_no_path.indexOf(".") != -1) {
                let filename_no_path_no_ext_ary = filename_no_path.split(".");
                let slide_ext = filename_no_path_no_ext_ary.pop();
                filename_no_path_no_ext = filename_no_path_no_ext_ary.join(".");
            }
            let slide_number = 1;
            if (filename_no_path_no_ext != "" && filename_no_path.indexOf("slide") != -1) {
                slide_number = Number(filename_no_path_no_ext.substr(5));
            }
            const slide = new Slide({
                fileName,
                index: i,
                processFullTheme: this.options.processFullTheme,
                mediaProcess: this.options.mediaProcess,
                resourceReader: this.resourceReader,
                basicInfo: this.basicInfo,
                tableStyles: this.tableStyles,
                styleTable: this.styleTable,
                slideFactor: this.options.slideFactor!,
                fontSizeFactor: this.options.fontSizeFactor!,
            });
            this.slides.push(slide);
            let slideHtml = await slide.processSingleSlide();
            post_ary.push({
                "type": "slide",
                "data": slideHtml,
                "slide_num": slide_number,
                "file_name": filename_no_path_no_ext
            });
        }

        post_ary.sort(function (a, b) {
            return a.slide_num - b.slide_num;
        });

        post_ary.push({
            "type": "globalCSS",
            "data": this.styleTable.genGlobalCSS()
        });

        this.htmlResultArray = post_ary;

        return post_ary;
    }

    private async readXmlFile(filename: string) {
        return this.resourceReader?.readXmlFile(filename);
    }

    private async getBasicInfo() {
        //get app version
        const [ContentTypesData, app, presentation] = await Promise.all([
            this.readXmlFile("[Content_Types].xml"),
            this.readXmlFile("docProps/app.xml"),
            this.readXmlFile("ppt/presentation.xml"),
        ]);
        let subObj = ContentTypesData["Types"]["Override"];
        let slidesLocArray = [];
        let slideLayoutsLocArray = [];
        for (var i = 0; i < subObj.length; i++) {
            switch (subObj[i]["attrs"]?.["ContentType"]) {
                case "application/vnd.openxmlformats-officedocument.presentationml.slide+xml":
                    slidesLocArray.push(subObj[i]["attrs"]?.["PartName"].substr(1));
                    break;
                case "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml":
                    slideLayoutsLocArray.push(subObj[i]["attrs"]?.["PartName"].substr(1));
                    break;
                default:
            }
        }
        let sldSzAttrs = presentation["p:presentation"]["p:sldSz"]["attrs"];
        let sldSzWidth = parseInt(sldSzAttrs["cx"]);
        let sldSzHeight = parseInt(sldSzAttrs["cy"]);
        let sldSzType = sldSzAttrs["type"];
        return {
            slides: slidesLocArray,
            slideLayouts: slideLayoutsLocArray,
            defaultTextStyle: presentation["p:presentation"]["p:defaultTextStyle"],
            width: sldSzWidth * this.options.slideFactor!,
            height: sldSzHeight * this.options.slideFactor!
        };
    }

    /*
    shapePolygon(sidesNum) {
        let sides  = sidesNum;
        let radius = 100;
        let angle  = 2 * Math.PI / sides;
        let points = []; 
        
        for (var i = 0; i < sides; i++) {
            points.push(radius + radius * Math.sin(i * angle));
            points.push(radius - radius * Math.cos(i * angle));
        }
        
        return points;
    }
    */

    async render(container: HTMLElement) {
        if (container) {
            this.options.container = container;
        }
        if (!this.options.container) {
            throw new Error("Container is not set");
        }
        await this.convertPromise;
        let htmlString = "";
        const result = this.htmlResultArray;
        for (var i = 0; i < result.length; i++) {
            switch (result[i]["type"]) {
                case "slide":
                    htmlString += (result[i]["data"]);
                    break;
                case "pptx-thumb":
                    //$("#pptx-thumb").attr("src", "data:image/jpeg;base64," +result[i]["data"]);
                    break;
                case "globalCSS":
                    //console.log(result[i]["data"])
                    htmlString += ("<style>" + result[i]["data"] + "</style>");
                    break;
                default:
            }
        }
        this.options.container.style.height = (this.basicInfo?.height ?? 0) * (this.basicInfo?.slides.length ?? 0) + "px";
        this.options.container.innerHTML = htmlString;
        this.postRender();
    }

    private setNumericBullets(elem: NodeListOf<Element>) {
        let prgrphs_arry = elem;
        for (var i = 0; i < prgrphs_arry.length; i++) {
            let buSpan = prgrphs_arry[i].querySelectorAll('.numeric-bullet-style');
            if (buSpan.length > 0) {
                //console.log("DIV-"+i+":");
                let prevBultTyp = "";
                let prevBultLvl = "";
                let buletIndex = 0;
                let tmpArry = new Array();
                let tmpArryIndx = 0;
                let buletTypSrry = new Array();
                for (var j = 0; j < buSpan.length; j++) {
                    let bult_typ = buSpan[j].getAttribute("data-bulltname") ?? '';
                    let bult_lvl = buSpan[j].getAttribute("data-bulltlvl") ?? '';
                    //console.log(j+" - "+bult_typ+" lvl: "+bult_lvl );
                    if (buletIndex == 0) {
                        prevBultTyp = bult_typ;
                        prevBultLvl = bult_lvl;
                        tmpArry[tmpArryIndx] = buletIndex;
                        buletTypSrry[tmpArryIndx] = bult_typ;
                        buletIndex++;
                    } else {
                        if (bult_typ == prevBultTyp && bult_lvl == prevBultLvl) {
                            prevBultTyp = bult_typ;
                            prevBultLvl = bult_lvl;
                            buletIndex++;
                            tmpArry[tmpArryIndx] = buletIndex;
                            buletTypSrry[tmpArryIndx] = bult_typ;
                        } else if (bult_typ != prevBultTyp && bult_lvl == prevBultLvl) {
                            prevBultTyp = bult_typ;
                            prevBultLvl = bult_lvl;
                            tmpArryIndx++;
                            tmpArry[tmpArryIndx] = buletIndex;
                            buletTypSrry[tmpArryIndx] = bult_typ;
                            buletIndex = 1;
                        } else if (bult_typ != prevBultTyp && Number(bult_lvl) > Number(prevBultLvl)) {
                            prevBultTyp = bult_typ;
                            prevBultLvl = bult_lvl;
                            tmpArryIndx++;
                            tmpArry[tmpArryIndx] = buletIndex;
                            buletTypSrry[tmpArryIndx] = bult_typ;
                            buletIndex = 1;
                        } else if (bult_typ != prevBultTyp && Number(bult_lvl) < Number(prevBultLvl)) {
                            prevBultTyp = bult_typ;
                            prevBultLvl = bult_lvl;
                            tmpArryIndx--;
                            buletIndex = tmpArry[tmpArryIndx] + 1;
                        }
                    }
                    //console.log(buletTypSrry[tmpArryIndx]+" - "+buletIndex);
                    let numIdx = getNumTypeNum(buletTypSrry[tmpArryIndx], buletIndex);
                    buSpan[j].innerHTML = numIdx;
                }
            }
        }
    }

    // d3 and nv d3 are initialized after the PPTX is rendered
    // so we need to call this function after the PPTX is rendered
    postRender() {
        if (this.postRenderDone) return;
        this.slides.forEach(slide => {
            slide.postRender();
        })
        this.setNumericBullets(document.querySelectorAll(".block"));
        this.setNumericBullets(document.querySelectorAll("table td"));
        this.postRenderDone = true;
    }
}
