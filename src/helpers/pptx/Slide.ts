// PPTX.ts is a rewrite of the original pptxjs library.

import { v4 as uuidv4 } from 'uuid';
import _ from 'lodash';
import tinycolor from 'tinycolor2';
import './assets/pptxjs.less';
import { escapeHtml, estimateWidth, getHtmlBullet } from './utils/html';
import { getBase64ImageDimensions} from './utils/image';
import { getLinerGrandient, getSvgGradient } from './utils/svg';
import { genTypedShape } from './utils/shape/genTypedShape';
import { extractFileExtension, getMimeType, isVideoLink } from './utils/file';
import { angleToDegrees, applyHueMod, applySatMod, applyTint, applyShade, applyLumOff, applyLumMod, rtlLangs, toHex, colorMap, hslToRgb } from './utils/color';
import './assets/d3.min.js';
import './assets/nv.d3.min.js';
import { StyleTable } from './Style';
import { eachElement, getTextByPathList, setTextByPathList } from './utils/object';
import { genCustShape } from './utils/shape/genCustShape';
import { ResourceReader } from './ResourceReader';

interface SlideOptions {
    fileName?: string;
    index?: number;
    processFullTheme?: boolean | string;
    mediaProcess?: boolean;
    resourceReader?: ResourceReader | null;
    slideFactor: number;
    fontSizeFactor: number;
    basicInfo?: {
        slides: string[];
        slideLayouts: string[];
        defaultTextStyle: any;
        width: number;
        height: number;
    } | null;
    tableStyles?: any;
    styleTable?: StyleTable;
}

interface Resources {
    slideLayoutContent: any,
    slideLayoutTables: any;
    slideMasterContent: any;
    slideMasterTables: any;
    slideContent: any;
    slideResObj: any;
    slideMasterTextStyles: any;
    layoutResObj: any;
    masterResObj: any;
    themeContent: any;
    themeResObj: any;
    digramFileContent: any;
    diagramResObj: any;
    tableStyles?: any;
}

export class Slide {
    private options: SlideOptions = {
        processFullTheme: true,
        mediaProcess: false,
        slideFactor: 96 / 914400,
        fontSizeFactor: 4 / 3.2,
    }

    private isFirstBr = false;
    private MsgQueue: any[] = [];
    private postRenderDone = false;

    private resources?: Resources;

    constructor(options: SlideOptions) {
        _.assign(this.options, options);
    }

    private async readXmlFile(filename: string) {
        return this.options.resourceReader?.readXmlFile(filename);
    }

    async processSingleSlide() {
        const sldFileName = this.options.fileName;
        const index = this.options.index;
        if (!sldFileName || !index) return '';
        let resName = sldFileName.replace("slides/slide", "slides/_rels/slide") + ".rels";
        let resContent = await this.readXmlFile(resName);
        let RelationshipArray = resContent["Relationships"]["Relationship"];
        let layoutFilename = "";
        let diagramFilename = "";
        let slideResObj: Record<string, any> = {};
        if (RelationshipArray.constructor === Array) {
            for (var i = 0; i < RelationshipArray.length; i++) {
                switch (RelationshipArray[i]["attrs"]?.["Type"]) {
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout":
                        layoutFilename = RelationshipArray[i]["attrs"]?.["Target"].replace("../", "ppt/");
                        break;
                    case "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing":
                        diagramFilename = RelationshipArray[i]["attrs"]?.["Target"].replace("../", "ppt/");
                        slideResObj[RelationshipArray[i]["attrs"]?.["Id"]] = {
                            "type": RelationshipArray[i]["attrs"]?.["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            "target": RelationshipArray[i]["attrs"]?.["Target"].replace("../", "ppt/")
                        };
                        break;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide":
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart":
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink":
                    default:
                        slideResObj[RelationshipArray[i]["attrs"]?.["Id"]] = {
                            "type": RelationshipArray[i]["attrs"]?.["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            "target": RelationshipArray[i]["attrs"]?.["Target"].replace("../", "ppt/")
                        };
                }
            }
        } else {
            layoutFilename = RelationshipArray["attrs"]?.["Target"].replace("../", "ppt/");
        }
        //console.log(slideResObj);
        // Open slideLayoutXX.xml
        let slideLayoutContent = await this.readXmlFile(layoutFilename);
        let slideLayoutTables = this.indexNodes(slideLayoutContent);
        let sldLayoutClrOvr = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping"]);

        //console.log(slideLayoutClrOvride);
        if (sldLayoutClrOvr !== undefined) {
            // slideLayoutClrOvride = sldLayoutClrOvr["attrs"];
        }
        // =====< Step 2 >=====
        // Read slide master filename of the slidelayout (Get slideMasterXX.xml)
        // @resName: ppt/slideLayouts/slideLayout1.xml
        // @masterName: ppt/slideLayouts/_rels/slideLayout1.xml.rels
        let slideLayoutResFilename = layoutFilename.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout") + ".rels";
        let slideLayoutResContent = await this.readXmlFile(slideLayoutResFilename);
        RelationshipArray = slideLayoutResContent["Relationships"]["Relationship"];
        let masterFilename = "";
        let layoutResObj: Record<string, any> = {};
        if (RelationshipArray.constructor === Array) {
            for (var i = 0; i < RelationshipArray.length; i++) {
                switch (RelationshipArray[i]["attrs"]?.["Type"]) {
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster":
                        masterFilename = RelationshipArray[i]["attrs"]?.["Target"].replace("../", "ppt/");
                        break;
                    default:
                        layoutResObj[RelationshipArray[i]["attrs"]?.["Id"]] = {
                            "type": RelationshipArray[i]["attrs"]?.["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            "target": RelationshipArray[i]["attrs"]?.["Target"].replace("../", "ppt/")
                        };
                }
            }
        } else {
            masterFilename = RelationshipArray["attrs"]?.["Target"].replace("../", "ppt/");
        }

        let slideMasterResFilename = masterFilename.replace("slideMasters/slideMaster", "slideMasters/_rels/slideMaster") + ".rels";
        const [slideContent, slideMasterContent, slideMasterResContent] = await Promise.all([
            this.readXmlFile(sldFileName),
            // Open slideMasterXX.xml
            this.readXmlFile(masterFilename),
            //Open slideMasterXX.xml.rels
            this.readXmlFile(slideMasterResFilename),
        ])
        let slideMasterTextStyles = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:txStyles"]);
        let slideMasterTables = this.indexNodes(slideMasterContent);

        RelationshipArray = slideMasterResContent["Relationships"]["Relationship"];
        let themeFilename = "";
        let masterResObj: Record<string, any> = {};
        if (RelationshipArray.constructor === Array) {
            for (var i = 0; i < RelationshipArray.length; i++) {
                switch (RelationshipArray[i]["attrs"]?.["Type"]) {
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme":
                        themeFilename = RelationshipArray[i]["attrs"]?.["Target"].replace("../", "ppt/");
                        break;
                    default:
                        masterResObj[RelationshipArray[i]["attrs"]?.["Id"]] = {
                            "type": RelationshipArray[i]["attrs"]?.["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            "target": RelationshipArray[i]["attrs"]?.["Target"].replace("../", "ppt/")
                        };
                }
            }
        } else {
            themeFilename = RelationshipArray["attrs"]?.["Target"].replace("../", "ppt/");
        }
        //console.log(themeFilename)
        //Load Theme file
        let themeResObj: Record<string, any> = {};
        let themeContent = null;
        if (themeFilename !== undefined) {
            let themeName = themeFilename.split("/").pop() as string;
            let themeResFileName = themeFilename.replace(themeName, "_rels/" + themeName) + ".rels";
            //console.log("themeFilename: ", themeFilename, ", themeName: ", themeName, ", themeResFileName: ", themeResFileName)
            themeContent = await this.readXmlFile(themeFilename);
            let themeResContent = await this.readXmlFile(themeResFileName);
            if (themeResContent !== null) {
                let relationshipArray = themeResContent["Relationships"]["Relationship"];
                if (relationshipArray !== undefined) {
                    let themeFilename = "";
                    if (relationshipArray.constructor === Array) {
                        for (var i = 0; i < relationshipArray.length; i++) {
                            themeResObj[relationshipArray[i]["attrs"]?.["Id"]] = {
                                "type": relationshipArray[i]["attrs"]?.["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                                "target": relationshipArray[i]["attrs"]?.["Target"].replace("../", "ppt/")
                            };
                        }
                    } else {
                        //console.log("theme relationshipArray : ", relationshipArray)
                        themeResObj[relationshipArray["attrs"]?.["Id"]] = {
                            "type": relationshipArray["attrs"]?.["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            "target": relationshipArray["attrs"]?.["Target"].replace("../", "ppt/")
                        };
                    }
                }
            }
        }
        //Load diagram file
        let diagramResObj: Record<string, any> = {};
        let digramFileContent = {};
        if (diagramFilename !== undefined) {
            let diagName = diagramFilename.split("/").pop() as string;
            let diagramResFileName = diagramFilename.replace(diagName, "_rels/" + diagName) + ".rels";
            //console.log("diagramFilename: ", diagramFilename, ", themeName: ", themeName, ", diagramResFileName: ", diagramResFileName)
            digramFileContent = await this.readXmlFile(diagramFilename);
            if (digramFileContent !== null && digramFileContent !== undefined && digramFileContent != "") {
                let digramFileContentObjToStr = JSON.stringify(digramFileContent);
                digramFileContentObjToStr = digramFileContentObjToStr.replace(/dsp:/g, "p:");
                digramFileContent = JSON.parse(digramFileContentObjToStr);
            }

            let digramResContent = await this.readXmlFile(diagramResFileName);
            if (digramResContent !== null) {
                let relationshipArray = digramResContent["Relationships"]["Relationship"];
                let themeFilename = "";
                if (relationshipArray.constructor === Array) {
                    for (var i = 0; i < relationshipArray.length; i++) {
                        diagramResObj[relationshipArray[i]["attrs"]?.["Id"]] = {
                            "type": relationshipArray[i]["attrs"]?.["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            "target": relationshipArray[i]["attrs"]?.["Target"].replace("../", "ppt/")
                        };
                    }
                } else {
                    //console.log("theme relationshipArray : ", relationshipArray)
                    diagramResObj[relationshipArray["attrs"]?.["Id"]] = {
                        "type": relationshipArray["attrs"]?.["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        "target": relationshipArray["attrs"]?.["Target"].replace("../", "ppt/")
                    };
                }
            }
        }
        //console.log("diagramResObj: " , diagramResObj)
        // =====< Step 3 >=====
        let nodes = slideContent["p:sld"]["p:cSld"]["p:spTree"];
        this.resources = {
            "slideLayoutContent": slideLayoutContent,
            "slideLayoutTables": slideLayoutTables,
            "slideMasterContent": slideMasterContent,
            "slideMasterTables": slideMasterTables,
            "slideContent": slideContent,
            "slideResObj": slideResObj,
            "slideMasterTextStyles": slideMasterTextStyles,
            "layoutResObj": layoutResObj,
            "masterResObj": masterResObj,
            "themeContent": themeContent,
            "themeResObj": themeResObj,
            "digramFileContent": digramFileContent,
            "diagramResObj": diagramResObj,
        };
        let bgResult = "";
        if (this.options.processFullTheme === true) {
            bgResult = await this.getBackground();
        }

        let bgColor: string | undefined = "";
        if (this.options.processFullTheme == "colorsAndImageOnly") {
            bgColor = await this.getSlideBackgroundFill();
        }

        let result = "<div class='slide' style='width:" + this.options.basicInfo?.width + "px; height:" + this.options.basicInfo?.height + "px;" + bgColor + "'>"
        result += bgResult;
        for (var nodeKey in nodes) {
            if (nodes[nodeKey].constructor === Array) {
                for (var i = 0; i < nodes[nodeKey].length; i++) {
                    result += await this.processNodesInSlide(nodeKey, nodes[nodeKey][i], nodes, "slide");
                }
            } else {
                result += await this.processNodesInSlide(nodeKey, nodes[nodeKey], nodes, "slide");
            }
        }
        return result + "</div>";
    }

    private indexNodes(content: any) {

        let keys = Object.keys(content);
        let spTreeNode = content[keys[0]]["p:cSld"]["p:spTree"];

        let idTable: Record<string, any> = {};
        let idxTable: Record<string, any> = {};
        let typeTable: Record<string, any> = {};

        for (var key in spTreeNode) {

            if (key == "p:nvGrpSpPr" || key == "p:grpSpPr") {
                continue;
            }

            let targetNode = spTreeNode[key];

            if (targetNode.constructor === Array) {
                for (var i = 0; i < targetNode.length; i++) {
                    let nvSpPrNode = targetNode[i]["p:nvSpPr"];
                    let id = getTextByPathList(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
                    let idx = getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "idx"]);
                    let type = getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "type"]);

                    if (id !== undefined) {
                        idTable[id] = targetNode[i];
                    }
                    if (idx !== undefined) {
                        idxTable[idx] = targetNode[i];
                    }
                    if (type !== undefined) {
                        typeTable[type] = targetNode[i];
                    }
                }
            } else {
                let nvSpPrNode = targetNode["p:nvSpPr"];
                let id = getTextByPathList(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
                let idx = getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "idx"]);
                let type = getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "type"]);

                if (id !== undefined) {
                    idTable[id] = targetNode;
                }
                if (idx !== undefined) {
                    idxTable[idx] = targetNode;
                }
                if (type !== undefined) {
                    typeTable[type] = targetNode;
                }
            }

        }

        return { "idTable": idTable, "idxTable": idxTable, "typeTable": typeTable };
    }

    private async processNodesInSlide(nodeKey: any, nodeValue: any, nodes: any, source: any, sType?: any) {
        let result = "";

        switch (nodeKey) {
            case "p:sp":    // Shape, Text
                result = await this.processSpNode(nodeValue, nodes, source, sType);
                break;
            case "p:cxnSp":    // Shape, Text (with connection)
                result = await this.processCxnSpNode(nodeValue, nodes, source, sType);
                break;
            case "p:pic":    // Picture
                result = await this.processPicNode(nodeValue, nodes, source, sType);
                break;
            case "p:graphicFrame":    // Chart, Diagram, Table
                result = await this.processGraphicFrameNode(nodeValue, source, sType);
                break;
            case "p:grpSp":
                result = await this.processGroupSpNode(nodeValue, source);
                break;
            case "mc:AlternateContent": //Equations and formulas as Image
                let mcFallbackNode = getTextByPathList(nodeValue, ["mc:Fallback"]);
                result = await this.processGroupSpNode(mcFallbackNode, source);
                break;
            default:
        }

        return result;

    }

    private async processGroupSpNode(node: any, source: any) {
        //console.log("processGroupSpNode: node: ", node)
        let xfrmNode = getTextByPathList(node, ["p:grpSpPr", "a:xfrm"]);
        let top, left, width, height, sType;
        let transform = "";
        if (xfrmNode !== undefined) {
            let x = parseInt(xfrmNode["a:off"]["attrs"]?.["x"]) * this.options.slideFactor;
            let y = parseInt(xfrmNode["a:off"]["attrs"]?.["y"]) * this.options.slideFactor;
            let chx = parseInt(xfrmNode["a:chOff"]["attrs"]?.["x"]) * this.options.slideFactor;
            let chy = parseInt(xfrmNode["a:chOff"]["attrs"]?.["y"]) * this.options.slideFactor;
            let cx = parseInt(xfrmNode["a:ext"]["attrs"]?.["cx"]) * this.options.slideFactor;
            let cy = parseInt(xfrmNode["a:ext"]["attrs"]?.["cy"]) * this.options.slideFactor;
            let chcx = parseInt(xfrmNode["a:chExt"]["attrs"]?.["cx"]) * this.options.slideFactor;
            let chcy = parseInt(xfrmNode["a:chExt"]["attrs"]?.["cy"]) * this.options.slideFactor;
            let rotate = parseInt(xfrmNode["attrs"]?.["rot"])
            top = chy,
            left = chx,
            width = chcx,
            height = chcy;

            sType = "group";
            if (x !== chx) {
                transform += " translateX(" +  (x - chx)+ "px)";
            }
            if (y !== chy) {
                transform += " translateY(" +  (y - chy)+ "px)";
            }
            if (cx !== chcx) {
                transform += " scaleX(" + cx / chcx + ")";
            }
            if (cy !== chcy) {
                transform += " scaleY(" + cy / chcy + ")";
            }
            if (!isNaN(rotate)) {
                rotate = angleToDegrees(rotate);
                transform += "rotate(" + rotate + "deg) ";
            }
        }
        let grpStyle = "";

        if (transform !== "") {
            grpStyle += "transform: " + transform + "; transform-origin: top left;";
        }

        if (top !== undefined) {
            grpStyle += "top: " + top + "px;";
        }
        if (left !== undefined) {
            grpStyle += "left: " + left + "px;";
        }
        if (width !== undefined) {
            grpStyle += "width:" + width + "px;";
        }
        if (height !== undefined) {
            grpStyle += "height: " + height + "px;";
        }

        let id = node["p:nvGrpSpPr"]?.["p:cNvPr"]?.["attrs"]?.["id"];
        let order = node["attrs"]?.["order"] || _.toInteger(id);

        let result = "<div class='block group' style='z-index: " + order + ";" + grpStyle + "'>";
        // wrap the group children with a div to abtain the correct position of the group and its children
        result += "<div style='top: -" + top + "px; left: -" + left + "px; position: absolute;'>";

        // Procsee all child nodes
        for (var nodeKey in node) {
            if (node[nodeKey].constructor === Array) {
                for (var i = 0; i < node[nodeKey].length; i++) {
                    result += await this.processNodesInSlide(nodeKey, node[nodeKey][i], node, source, sType);
                }
            } else {
                result += await this.processNodesInSlide(nodeKey, node[nodeKey], node, source, sType);
            }
        }

        result += "</div>";
        result += "</div>";

        return result;
    }

    private async processSpNode(node: any, pNode: any, source: any, sType: any) {

        /*
        *  958    <xsd:complexType name="CT_GvmlShape">
        *  959   <xsd:sequence>
        *  960     <xsd:element name="nvSpPr" type="CT_GvmlShapeNonVisual"     minOccurs="1" maxOccurs="1"/>
        *  961     <xsd:element name="spPr"   type="CT_ShapeProperties"        minOccurs="1" maxOccurs="1"/>
        *  962     <xsd:element name="txSp"   type="CT_GvmlTextShape"          minOccurs="0" maxOccurs="1"/>
        *  963     <xsd:element name="style"  type="CT_ShapeStyle"             minOccurs="0" maxOccurs="1"/>
        *  964     <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
        *  965   </xsd:sequence>
        *  966 </xsd:complexType>
        */

        let id = getTextByPathList(node, ["p:nvSpPr", "p:cNvPr", "attrs", "id"]);
        let name = getTextByPathList(node, ["p:nvSpPr", "p:cNvPr", "attrs", "name"]);
        let idx = (getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "idx"]) === undefined) ? undefined : getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "idx"]);
        let type = (getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]) === undefined) ? undefined : getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
        let order = getTextByPathList(node, ["attrs", "order"]) || _.toInteger(id);
        let isUserDrawnBg;
        if (source == "slideLayoutBg" || source == "slideMasterBg") {
            let userDrawn = getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "attrs", "userDrawn"]);
            if (userDrawn == "1") {
                isUserDrawnBg = true;
            } else {
                isUserDrawnBg = false;
            }
        }
        let slideLayoutSpNode = undefined;
        let slideMasterSpNode = undefined;

        if (idx !== undefined) {
            slideLayoutSpNode = this.resources?.slideLayoutTables["idxTable"][idx];
            if (type !== undefined) {
                slideMasterSpNode = this.resources?.slideMasterTables["typeTable"][type];
            } else {
                slideMasterSpNode = this.resources?.slideMasterTables["idxTable"][idx];
            }
        } else {
            if (type !== undefined) {
                slideLayoutSpNode = this.resources?.slideLayoutTables["typeTable"][type];
                slideMasterSpNode = this.resources?.slideMasterTables["typeTable"][type];
            }
        }

        if (type === undefined) {
            const txBoxVal = getTextByPathList(node, ["p:nvSpPr", "p:cNvSpPr", "attrs", "txBox"]);
            if (txBoxVal == "1") {
                type = "textBox";
            }
        }
        if (type === undefined) {
            type = getTextByPathList(slideLayoutSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
            if (type === undefined) {
                //type = getTextByPathList(slideMasterSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                if (source == "diagramBg") {
                    type = "diagram";
                } else {

                    type = "obj"; //default type
                }
            }
        }
        //console.log("processSpNode type:", type, "idx:", idx);
        return await this.genShape(node, pNode, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, isUserDrawnBg, sType, source);
    }

    private async processCxnSpNode(node: any, pNode: any, source: any, sType: any) {

        let id = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]?.["id"];
        let name = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]?.["name"];
        let idx = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]?.["idx"];
        let type = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]?.["type"];
        //<p:cNvCxnSpPr>(<p:cNvCxnSpPr>, <a:endCxn>)
        let order = node["attrs"]?.["order"] || _.toInteger(id);

        return await this.genShape(node, pNode, undefined, undefined, id, name, idx, type, order, undefined, sType, source);
    }

    private async genShape(node: any, pNode: any, slideLayoutSpNode: any, slideMasterSpNode: any, id: any, name: any, idx: any, type: any, order: any, isUserDrawnBg: any, sType: any, source: any) {
        //var dltX = 0;
        //var dltY = 0;
        let xfrmList = ["p:spPr", "a:xfrm"];
        let slideXfrmNode = getTextByPathList(node, xfrmList);
        let slideLayoutXfrmNode = getTextByPathList(slideLayoutSpNode, xfrmList);
        let slideMasterXfrmNode = getTextByPathList(slideMasterSpNode, xfrmList);

        let result = "";
        let shpId = getTextByPathList(node, ["attrs", "order"]);
        //console.log("shpId: ",shpId)
        let shapType = getTextByPathList(node, ["p:spPr", "a:prstGeom", "attrs", "prst"]);

        //custGeom - Amir
        let custShapType = getTextByPathList(node, ["p:spPr", "a:custGeom"]);

        let isFlipV = false;
        let isFlipH = false;
        let flip = "";
        if (getTextByPathList(slideXfrmNode, ["attrs", "flipV"]) === "1") {
            isFlipV = true;
        }
        if (getTextByPathList(slideXfrmNode, ["attrs", "flipH"]) === "1") {
            isFlipH = true;
        }
        if (isFlipH && !isFlipV) {
            flip = " scale(-1,1)"
        } else if (!isFlipH && isFlipV) {
            flip = " scale(1,-1)"
        } else if (isFlipH && isFlipV) {
            flip = " scale(-1,-1)"
        }
        /////////////////////////Amir////////////////////////
        //rotate
        let rotate = angleToDegrees(getTextByPathList(slideXfrmNode, ["attrs", "rot"]));

        //console.log("genShape rotate: " + rotate);
        let txtRotate;
        let txtXframeNode = getTextByPathList(node, ["p:txXfrm"]);
        if (txtXframeNode !== undefined) {
            let txtXframeRot = getTextByPathList(txtXframeNode, ["attrs", "rot"]);
            if (txtXframeRot !== undefined) {
                txtRotate = angleToDegrees(txtXframeRot) + 90;
            }
        } else {
            txtRotate = rotate;
        }
        if (shapType == "gear9") {
            txtRotate = 0;
        }
        //////////////////////////////////////////////////
        let border: any;
        let off, x, y, ext, w = 0, h = 0, grndFillFlg, imgFillFlg, clrFillType, fillColor, headEndNodeAttrs, tailEndNodeAttrs;
        let oShadowSvgUrlStr = ""
        if (shapType !== undefined || custShapType !== undefined /*&& slideXfrmNode !== undefined*/) {
            off = getTextByPathList(slideXfrmNode, ["a:off", "attrs"]);
            x = parseInt(off["x"]) * this.options.slideFactor;
            y = parseInt(off["y"]) * this.options.slideFactor;

            ext = getTextByPathList(slideXfrmNode, ["a:ext", "attrs"]);
            w = parseInt(ext["cx"]) * this.options.slideFactor;
            h = parseInt(ext["cy"]) * this.options.slideFactor;

            let svgCssName = "_svg_css_" + this.options.styleTable?.nextIndex() + "_" + Math.floor(Math.random() * 1001);
            //console.log("name:", name, "svgCssName: ", svgCssName)
            let effectsClassName = svgCssName + "_effects";
            result += "<svg class='drawing " + svgCssName + " " + effectsClassName + " ' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name + "'" +
                "' style='" +
                this.getPosition(slideXfrmNode, pNode, undefined, undefined, sType) +
                this.getSize(slideXfrmNode, pNode, undefined, undefined, sType) +
                " min-height: 1px;" +
                " z-index: " + order + ";" +
                "transform: rotate(" + ((rotate !== undefined) ? rotate : 0) + "deg)" + flip + ";" +
                "'>";
            result += '<defs>'
            // Fill Color
            fillColor = await this.getShapeFill(node, pNode, true, source);
            //console.log("genShape: fillColor: ", fillColor)
            grndFillFlg = false;
            imgFillFlg = false;
            clrFillType = this.getFillType(getTextByPathList(node, ["p:spPr"]));
            if (clrFillType == "GROUP_FILL") {
                clrFillType = this.getFillType(getTextByPathList(pNode, ["p:grpSpPr"]));
            }
            // if (clrFillType == "") {
            //     let clrFillType = getFillType(getTextByPathList(node, ["p:style","a:fillRef"]));
            // }
            //console.log("genShape: fillColor: ", fillColor, ", clrFillType: ", clrFillType, ", node: ", node)
            /////////////////////////////////////////                    
            if (clrFillType == "GRADIENT_FILL") {
                grndFillFlg = true;
                let color_arry = fillColor.color;
                let angl = fillColor.rot + 90;
                let svgGrdnt = getSvgGradient(w, h, angl, color_arry, shpId);
                //fill="url(#linGrd)"
                //console.log("genShape: svgGrdnt: ", svgGrdnt)
                result += svgGrdnt;

            } else if (clrFillType == "PIC_FILL") {
                imgFillFlg = true;
                let svgBgImg = this.getSvgImagePattern(node, fillColor, shpId);
                //fill="url(#imgPtrn)"
                //console.log(svgBgImg)
                result += svgBgImg;
            } else if (clrFillType == "PATTERN_FILL") {
                let styleText = fillColor;
                if (this.options.styleTable?.has(styleText)) {
                    styleText += "do-nothing: " + svgCssName + ";";
                }
                this.options.styleTable?.update(styleText, {
                    "name": svgCssName,
                    "text": styleText
                });
                //}
                fillColor = "none";
            } else {
                if (clrFillType != "SOLID_FILL" && clrFillType != "PATTERN_FILL" &&
                    (shapType == "arc" ||
                        shapType == "bracketPair" ||
                        shapType == "bracePair" ||
                        shapType == "leftBracket" ||
                        shapType == "leftBrace" ||
                        shapType == "rightBrace" ||
                        shapType == "rightBracket")) { //Temp. solution  - TODO
                    fillColor = "none";
                }
            }
            // Border Color
            border = this.getBorder(node, pNode, true, "shape");

            headEndNodeAttrs = getTextByPathList(node, ["p:spPr", "a:ln", "a:headEnd", "attrs"]);
            tailEndNodeAttrs = getTextByPathList(node, ["p:spPr", "a:ln", "a:tailEnd", "attrs"]);
            // type: none, triangle, stealth, diamond, oval, arrow

            ////////////////////effects/////////////////////////////////////////////////////
            //p:spPr => a:effectLst =>
            //"a:blur"
            //"a:fillOverlay"
            //"a:glow"
            //"a:innerShdw"
            //"a:outerShdw"
            //"a:prstShdw"
            //"a:reflection"
            //"a:softEdge"
            //p:spPr => a:scene3d
            //"a:camera"
            //"a:lightRig"
            //"a:backdrop"
            //"a:extLst"?
            //p:spPr => a:sp3d
            //"a:bevelT"
            //"a:bevelB"
            //"a:extrusionClr"
            //"a:contourClr"
            //"a:extLst"?
            //////////////////////////////outerShdw///////////////////////////////////////////
            //not support sizing the shadow
            let outerShdwNode = getTextByPathList(node, ["p:spPr", "a:effectLst", "a:outerShdw"]);
            if (outerShdwNode !== undefined) {
                let chdwClrNode = this.getSolidFill(outerShdwNode);
                let outerShdwAttrs = outerShdwNode["attrs"] ?? {};

                //var algn = outerShdwAttrs["algn"];
                let dir = (outerShdwAttrs["dir"]) ? (parseInt(outerShdwAttrs["dir"]) / 60000) : 0;
                let dist = parseInt(outerShdwAttrs["dist"]) * this.options.slideFactor;//(px) //* (3 / 4); //(pt)
                //var rotWithShape = outerShdwAttrs["rotWithShape"];
                let blurRad = (outerShdwAttrs["blurRad"]) ? (parseInt(outerShdwAttrs["blurRad"]) * this.options.slideFactor) : ""; //+ "px"
                //var sx = (outerShdwAttrs["sx"]) ? (parseInt(outerShdwAttrs["sx"]) / 100000) : 1;
                //var sy = (outerShdwAttrs["sy"]) ? (parseInt(outerShdwAttrs["sy"]) / 100000) : 1;
                let vx = dist * Math.sin(dir * Math.PI / 180);
                let hx = dist * Math.cos(dir * Math.PI / 180);
                //SVG
                //var oShadowId = "outerhadow_" + shpId;
                //oShadowSvgUrlStr = "filter='url(#" + oShadowId+")'";
                //var shadowFilterStr = '<filter id="' + oShadowId + '" x="0" y="0" width="' + w * (6 / 8) + '" height="' + h + '">';
                //1:
                //shadowFilterStr += '<feDropShadow dx="' + vx + '" dy="' + hx + '" stdDeviation="' + blurRad * (3 / 4) + '" flood-color="#' + chdwClrNode +'" flood-opacity="1" />'
                //2:
                //shadowFilterStr += '<feFlood result="floodColor" flood-color="red" flood-opacity="0.5"   width="' + w * (6 / 8) + '" height="' + h + '"  />'; //#' + chdwClrNode +'
                //shadowFilterStr += '<feOffset result="offOut" in="SourceGraph ccfsdf-+ic"  dx="' + vx + '" dy="' + hx + '"/>'; //how much to offset
                //shadowFilterStr += '<feGaussianBlur result="blurOut" in="offOut" stdDeviation="' + blurRad*(3/4) +'"/>'; //tdDeviation is how much to blur
                //shadowFilterStr += '<feComponentTransfer><feFuncA type="linear" slope="0.5"/></feComponentTransfer>'; //slope is the opacity of the shadow
                //shadowFilterStr += '<feBlend in="SourceGraphic" in2="blurOut"  mode="normal" />'; //this contains the element that the filter is applied to
                //shadowFilterStr += '</filter>'; 
                //result += shadowFilterStr;

                //css:
                let svg_css_shadow = "filter:drop-shadow(" + hx + "px " + vx + "px " + blurRad + "px #" + chdwClrNode + ");";

                if (this.options.styleTable?.has(svg_css_shadow)) {
                    svg_css_shadow += "do-nothing: " + svgCssName + ";";
                }

                this.options.styleTable?.update(svg_css_shadow, {
                    "name": effectsClassName,
                    "text": svg_css_shadow
                });

            }
            ////////////////////////////////////////////////////////////////////////////////////////
            if ((headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) ||
                (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow"))) {
                let triangleMarker = "<marker id='markerTriangle_" + shpId + "' viewBox='0 0 10 10' refX='1' refY='5' markerWidth='5' markerHeight='5' stroke='" + border.color + "' fill='" + border.color +
                    "' orient='auto-start-reverse' markerUnits='strokeWidth'><path d='M 0 0 L 10 5 L 0 10 z' /></marker>";
                result += triangleMarker;
            }
            result += '</defs>'
        }
        if (shapType !== undefined && custShapType === undefined) {
            const fill = !imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")";
            const markerStart = headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow") ? "marker-start='url(#markerTriangle_" + shpId + ")' " : "";
            const markerEnd = tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow") ? "marker-end='url(#markerTriangle_" + shpId + ")' " : "";
            result += genTypedShape(
                shapType,
                w,
                h,
                border,
                fill,
                this.options.slideFactor,
                getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]),
                oShadowSvgUrlStr,
                markerStart,
                markerEnd
            );

            result += "</svg>";

            result += "<div class='block " + this.getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) + //block content
                " " + this.getContentDir(node, type) +
                "' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
                "' style='" +
                this.getPosition(slideXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType) +
                this.getSize(slideXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType) +
                " z-index: " + order + ";" +
                "transform: rotate(" + ((txtRotate !== undefined) ? txtRotate : 0) + "deg);" +
                "'>";

            // TextBody
            if (node["p:txBody"] !== undefined && (isUserDrawnBg === undefined || isUserDrawnBg === true)) {
                if (type != "diagram" && type != "textBox") {
                    type = "shape";
                }
                result += await this.genTextBody(node["p:txBody"], node, slideLayoutSpNode, slideMasterSpNode, type, idx); //type='shape'
            }
            result += "</div>";
        } else if (custShapType !== undefined) {
            //custGeom here - Amir ///////////////////////////////////////////////////////
            //http://officeopenxml.com/drwSp-custGeom.php
            let pathLstNode = getTextByPathList(custShapType, ["a:pathLst"]);
            let pathNodes = getTextByPathList(pathLstNode, ["a:path"]);
            //var pathNode = getTextByPathList(pathLstNode, ["a:path", "attrs"]);
            let maxX = parseInt(pathNodes["attrs"]?.["w"]);// * this.options.slideFactor;
            let maxY = parseInt(pathNodes["attrs"]?.["h"]);// * this.options.slideFactor;
            let cX = (1 / maxX) * w;
            let cY = (1 / maxY) * h;
            //console.log("w = "+w+"\nh = "+h+"\nmaxX = "+maxX +"\nmaxY = " + maxY);
            //cheke if it is close shape

            //console.log("custShapType : ", custShapType, ", pathLstNode: ", pathLstNode, ", node: ", node);//, ", y:", y, ", w:", w, ", h:", h);

            let moveToNode = getTextByPathList(pathNodes, ["a:moveTo"]);
            let total_shapes = moveToNode.length;

            let lnToNodes = pathNodes["a:lnTo"]; //total a:pt : 1
            let cubicBezToNodes = pathNodes["a:cubicBezTo"]; //total a:pt : 3
            let arcToNodes = pathNodes["a:arcTo"]; //total a:pt : 0?1? ; attrs: ~4 ()
            let closeNode = getTextByPathList(pathNodes, ["a:close"]); //total a:pt : 0
            //quadBezTo //total a:pt : 2 - TODO
            //console.log("ia moveToNode array: ", Array.isArray(moveToNode))
            if (!Array.isArray(moveToNode)) {
                moveToNode = [moveToNode];
            }
            //console.log("ia moveToNode array: ", Array.isArray(moveToNode))

            const fill = (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")
            result += genCustShape(moveToNode, lnToNodes, cubicBezToNodes, arcToNodes, closeNode, cX, cY, fill, border);

            result += "</svg>";
            result += "<div class='block " + this.getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) + //block content 
                " " + this.getContentDir(node, type) +
                "' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
                "' style='" +
                this.getPosition(slideXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType) +
                this.getSize(slideXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType) +
                " z-index: " + order + ";" +
                "transform: rotate(" + ((txtRotate !== undefined) ? txtRotate : 0) + "deg);" +
                "'>";

            // TextBody
            if (node["p:txBody"] !== undefined && (isUserDrawnBg === undefined || isUserDrawnBg === true)) {
                if (type != "diagram" && type != "textBox") {
                    type = "shape";
                }
                result += await this.genTextBody(node["p:txBody"], node, slideLayoutSpNode, slideMasterSpNode, type, idx); //type=shape
            }
            result += "</div>";

            // result = "";
        } else {
            result += "<div class='block " + this.getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) + //block content 
                " " + this.getContentDir(node, type) +
                "' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
                "' style='" +
                this.getPosition(slideXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType) +
                this.getSize(slideXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType) +
                this.getBorder(node, pNode, false, "shape") +
                await this.getShapeFill(node, pNode, false, source) +
                " z-index: " + order + ";" +
                "transform: rotate(" + ((txtRotate !== undefined) ? txtRotate : 0) + "deg);" +
                "'>";

            // TextBody
            if (node["p:txBody"] !== undefined && (isUserDrawnBg === undefined || isUserDrawnBg === true)) {
                result += await this.genTextBody(node["p:txBody"], node, slideLayoutSpNode, slideMasterSpNode, type, idx);
            }
            result += "</div>";

        }
        //console.log("div block result:\n", result)
        return result;
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
    private async processPicNode(node: any, pNode: any, source: string, sType: string) {
        //console.log("processPicNode node:", node, "source:", source, "sType:", sType);
        let rtrnData = "";
        let mediaPicFlag = false;
        let id = getTextByPathList(node, ["p:nvPicPr", "p:cNvPr", "attrs", "id"]);
        let order = node["attrs"]?.["order"] || _.toInteger(id);

        let rid = node["p:blipFill"]["a:blip"]["attrs"]?.["r:embed"];
        let resObj;
        if (source == "slideMasterBg") {
            resObj = this.resources?.masterResObj;
        } else if (source == "slideLayoutBg") {
            resObj = this.resources?.layoutResObj;
        } else {
            //imgName = this.resources["slideResObj"][rid]["target"];
            resObj = this.resources?.slideResObj;
        }
        let imgName = resObj[rid]["target"];

        //console.log("processPicNode imgName:", imgName);
        let xfrmNode = node["p:spPr"]["a:xfrm"];
        if (xfrmNode === undefined) {
            let idx = getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "idx"]);
            let type = getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "type"]);
            if (idx !== undefined) {
                xfrmNode = getTextByPathList(this.resources?.slideLayoutTables, ["idxTable", idx, "p:spPr", "a:xfrm"]);
            }
        }
        ///////////////////////////////////////Amir//////////////////////////////
        let rotate = 0;
        let rotateNode = getTextByPathList(node, ["p:spPr", "a:xfrm", "attrs", "rot"]);
        if (rotateNode !== undefined) {
            rotate = angleToDegrees(rotateNode);
        }
        //video
        let vdoNode = getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:videoFile"]);
        let vdoRid, vdoFile, vdoFileExt, vdoMimeType, uInt8Array, blob, vdoBlob, mediaSupportFlag = false, isVdeoLink = false;
        let mediaProcess = this.options.mediaProcess;
        if (vdoNode !== undefined && mediaProcess) {
            vdoRid = vdoNode["attrs"]?.["r:link"];
            vdoFile = resObj[vdoRid]["target"];
            let checkIfLink = isVideoLink(vdoFile);
            if (checkIfLink) {
                vdoFile = escapeHtml(vdoFile);
                //vdoBlob = vdoFile;
                isVdeoLink = true;
                mediaSupportFlag = true;
                mediaPicFlag = true;
            } else {
                vdoFileExt = extractFileExtension(vdoFile);
                if (vdoFileExt == "mp4" || vdoFileExt == "webm" || vdoFileExt == "ogg") {
                    uInt8Array = await this.options.resourceReader?.getFile(vdoFile, "arraybuffer");
                    vdoMimeType = getMimeType(vdoFileExt);
                    blob = new Blob([uInt8Array!], {
                        type: vdoMimeType
                    });
                    vdoBlob = URL.createObjectURL(blob);
                    mediaSupportFlag = true;
                    mediaPicFlag = true;
                }
            }
        }
        //Audio
        let audioNode = getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:audioFile"]);
        let audioRid, audioFile, audioFileExt, audioMimeType, uInt8ArrayAudio, blobAudio, audioBlob;
        let audioPlayerFlag = false;
        let audioObjc;
        if (audioNode !== undefined && mediaProcess) {
            audioRid = audioNode["attrs"]?.["r:link"];
            audioFile = resObj[audioRid]["target"];
            audioFileExt = extractFileExtension(audioFile);
            if (audioFileExt == "mp3" || audioFileExt == "wav" || audioFileExt == "ogg") {
                uInt8ArrayAudio = await this.options.resourceReader?.getFile(audioFile, "arraybuffer");
                blobAudio = new Blob([uInt8ArrayAudio!]);
                audioBlob = URL.createObjectURL(blobAudio);
                let cx = parseInt(xfrmNode["a:ext"]["attrs"]?.["cx"]) * 20;
                let cy = xfrmNode["a:ext"]["attrs"]?.["cy"];
                let x = parseInt(xfrmNode["a:off"]["attrs"]?.["x"]) / 2.5;
                let y = xfrmNode["a:off"]["attrs"]?.["y"];
                audioObjc = {
                    "a:ext": {
                        "attrs": {
                            "cx": cx,
                            "cy": cy
                        }
                    },
                    "a:off": {
                        "attrs": {
                            "x": x,
                            "y": y

                        }
                    }
                }
                audioPlayerFlag = true;
                mediaSupportFlag = true;
                mediaPicFlag = true;
            }
        }
        //console.log(node)
        //////////////////////////////////////////////////////////////////////////
        rtrnData = "<div class='block content' style='" +
            ((mediaProcess && audioPlayerFlag) ? this.getPosition(audioObjc, pNode, undefined, undefined) : this.getPosition(xfrmNode, pNode, undefined, undefined)) +
            ((mediaProcess && audioPlayerFlag) ? this.getSize(audioObjc, pNode, undefined, undefined, sType) : this.getSize(xfrmNode, pNode, undefined, undefined)) +
            this.getBorder(node, pNode, false, "shape") +
            " z-index: " + order + ";" +
            "transform: rotate(" + rotate + "deg);'>";
        if ((vdoNode === undefined && audioNode === undefined) || !mediaProcess || !mediaSupportFlag) {
            rtrnData += "<img src='" + await this.options.resourceReader?.getImgBase64Data(imgName) + "' style='width: 100%; height: 100%'/>";
        } else if ((vdoNode !== undefined || audioNode !== undefined) && mediaProcess && mediaSupportFlag) {
            if (vdoNode !== undefined && !isVdeoLink) {
                rtrnData += "<video  src='" + vdoBlob + "' controls style='width: 100%; height: 100%'>Your browser does not support the video tag.</video>";
            } else if (vdoNode !== undefined && isVdeoLink) {
                rtrnData += "<iframe   src='" + vdoFile + "' controls style='width: 100%; height: 100%'></iframe >";
            }
            if (audioNode !== undefined) {
                rtrnData += '<audio id="audio_player" controls ><source src="' + audioBlob + '"></audio>';
                //'<button onclick="audio_player.play()">Play</button>'+
                //'<button onclick="audio_player.pause()">Pause</button>';
            }
        }
        if (!mediaSupportFlag && mediaPicFlag) {
            rtrnData += "<span style='color:red;font-size:40px;position: absolute;'>This media file Not supported by HTML5</span>";
        }
        if ((vdoNode !== undefined || audioNode !== undefined) && !mediaProcess && mediaSupportFlag) {
            console.log("Founded supported media file but media process disabled (mediaProcess=false)");
        }
        rtrnData += "</div>";
        //console.log(rtrnData)
        return rtrnData;
    }

    private async processGraphicFrameNode(node: any, source: string, sType?: string) {

        let result = "";
        let graphicTypeUri = getTextByPathList(node, ["a:graphic", "a:graphicData", "attrs", "uri"]);

        switch (graphicTypeUri) {
            case "http://schemas.openxmlformats.org/drawingml/2006/table":
                result = await this.genTable(node);
                break;
            case "http://schemas.openxmlformats.org/drawingml/2006/chart":
                result = await this.genChart(node);
                break;
            case "http://schemas.openxmlformats.org/drawingml/2006/diagram":
                result = await this.genDiagram(node, source, sType);
                break;
            case "http://schemas.openxmlformats.org/presentationml/2006/ole":
                //result = genDiagram(node, this.resources, source, sType);
                let oleObjNode = getTextByPathList(node, ["a:graphic", "a:graphicData", "mc:AlternateContent", "mc:Fallback", "p:oleObj"]);

                if (oleObjNode === undefined) {
                    oleObjNode = getTextByPathList(node, ["a:graphic", "a:graphicData", "p:oleObj"]);
                }
                //console.log("node:", node, "oleObjNode:", oleObjNode)
                if (oleObjNode !== undefined) {
                    result = await this.processGroupSpNode(oleObjNode, source);
                }
                break;
            default:
        }

        return result;
    }

    private processSpPrNode(node: any) {

        /*
        * 2241 <xsd:complexType name="CT_ShapeProperties">
        * 2242   <xsd:sequence>
        * 2243     <xsd:element name="xfrm" type="CT_Transform2D"  minOccurs="0" maxOccurs="1"/>
        * 2244     <xsd:group   ref="EG_Geometry"                  minOccurs="0" maxOccurs="1"/>
        * 2245     <xsd:group   ref="EG_FillProperties"            minOccurs="0" maxOccurs="1"/>
        * 2246     <xsd:element name="ln" type="CT_LineProperties" minOccurs="0" maxOccurs="1"/>
        * 2247     <xsd:group   ref="EG_EffectProperties"          minOccurs="0" maxOccurs="1"/>
        * 2248     <xsd:element name="scene3d" type="CT_Scene3D"   minOccurs="0" maxOccurs="1"/>
        * 2249     <xsd:element name="sp3d" type="CT_Shape3D"      minOccurs="0" maxOccurs="1"/>
        * 2250     <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
        * 2251   </xsd:sequence>
        * 2252   <xsd:attribute name="bwMode" type="ST_BlackWhiteMode" use="optional"/>
        * 2253 </xsd:complexType>
        */

        // TODO:
    }

    private async genTextBody(textBodyNode: any, spNode: any, slideLayoutSpNode: any, slideMasterSpNode: any, type: string | undefined, idx: number | undefined, tbl_col_width?: number) {
        let text = "";
        let slideMasterTextStyles = this.resources?.slideMasterTextStyles;

        if (textBodyNode === undefined) {
            return text;
        }
        //rtl : <p:txBody>
        //          <a:bodyPr wrap="square" rtlCol="1">

        let pFontStyle = getTextByPathList(spNode, ["p:style", "a:fontRef"]);
        //console.log("genTextBody spNode: ", getTextByPathList(spNode,["p:spPr","a:xfrm","a:ext"]));

        //var lstStyle = textBodyNode["a:lstStyle"];

        let apNode = textBodyNode["a:p"];
        if (apNode.constructor !== Array) {
            apNode = [apNode];
        }

        for (var i = 0; i < apNode.length; i++) {
            let pNode = apNode[i];
            let rNode = pNode["a:r"];
            let fldNode = pNode["a:fld"];
            let brNode = pNode["a:br"];
            if (rNode !== undefined) {
                rNode = (rNode.constructor === Array) ? rNode : [rNode];
            }
            if (rNode !== undefined && fldNode !== undefined) {
                fldNode = (fldNode.constructor === Array) ? fldNode : [fldNode];
                rNode = rNode.concat(fldNode)
            }
            if (rNode !== undefined && brNode !== undefined) {
                this.isFirstBr = true;
                brNode = (brNode.constructor === Array) ? brNode : [brNode];
                brNode.forEach(function (item: any) {
                    item.type = "br";
                });
                if (brNode.length > 1) {
                    brNode.shift();
                }
                rNode = rNode.concat(brNode)
                //console.log("single a:p  rNode:", rNode, "brNode:", brNode )
                rNode.sort(function (a: any, b: any) {
                    return a.attrs?.order - b.attrs?.order;
                });
                //console.log("sorted rNode:",rNode)
            }
            //rtlStr = "";//"dir='"+isRTL+"'";
            let styleText = "";
            let marginsVer = this.getVerticalMargins(pNode, textBodyNode, type, idx);
            if (marginsVer != "") {
                styleText = marginsVer;
            }
            if (type == "body" || type == "obj" || type == "shape") {
                styleText += "font-size: 0px;";
                //styleText += "line-height: 0;";
                styleText += "font-weight: 100;";
                styleText += "font-style: normal;";
            }
            let cssName = "";

            if (this.options.styleTable?.has(styleText)) {
                cssName = this.options.styleTable?.get(styleText).name;
            } else {
                cssName = "_css_" + this.options.styleTable?.nextIndex();
                this.options.styleTable?.update(styleText, {
                    "name": cssName,
                    "text": styleText
                });
            }
            //console.log("textBodyNode: ", textBodyNode["a:lstStyle"])
            let prg_width_node = getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cx"]);
            let prg_height_node = getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cy"]);
            let sld_prg_width = ((prg_width_node !== undefined) ? ("width:" + (parseInt(prg_width_node) * this.options.slideFactor) + "px;") : "width:inherit;");
            let sld_prg_height = ((prg_height_node !== undefined) ? ("height:" + (parseInt(prg_height_node) * this.options.slideFactor) + "px;") : "");
            let prg_dir = this.getPregraphDir(pNode, textBodyNode, idx, type);
            text += "<div style='display: flex;" + sld_prg_width + sld_prg_height + "' class='slide-prgrph " + this.getHorizontalAlign(pNode, textBodyNode, idx, type, prg_dir) + " " +
                prg_dir + " " + cssName + "' >";
            let buText_ary = await this.genBuChar(pNode, i, spNode, textBodyNode, pFontStyle, idx, type);
            let isBullate = buText_ary ? true : false;
            let bu_width = (buText_ary && buText_ary[1] !== undefined && buText_ary[1] !== null && isBullate) ? buText_ary[1] + buText_ary[2] : 0;
            text += buText_ary ? buText_ary[0] : "";
            //get text margin 
            let margin_ary = this.getPregraphMargn(pNode, idx, type, isBullate);
            let margin = margin_ary[0];
            let mrgin_val = margin_ary[1];
            if (prg_width_node === undefined && tbl_col_width !== undefined && prg_width_node != 0) {
                //sorce : table text
                prg_width_node = tbl_col_width;
            }

            let prgrph_text = "";
            //var prgr_txt_art = [];
            let total_text_len = 0;
            if (rNode === undefined && pNode !== undefined) {
                // without r
                let prgr_text = this.genSpanElement(pNode, undefined, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, 1, isBullate);
                if (isBullate) {
                    total_text_len += estimateWidth(prgr_text);
                }
                prgrph_text += prgr_text;
            } else if (rNode !== undefined) {
                // with multi r
                for (var j = 0; j < rNode.length; j++) {
                    let prgr_text = this.genSpanElement(rNode[j], j, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, rNode.length, isBullate);
                    if (isBullate) {
                        total_text_len += estimateWidth(prgr_text);
                    }
                    prgrph_text += prgr_text;
                }
            }

            prg_width_node = parseInt(prg_width_node) * this.options.slideFactor - (bu_width as number) + (mrgin_val as number);
            if (isBullate) {
                //get prg_width_node if there is a bulltes
                //console.log("total_text_len: ", total_text_len, "prg_width_node:", prg_width_node)
                if (total_text_len < prg_width_node) {
                    prg_width_node = total_text_len + (bu_width as number) + (mrgin_val as number);
                }
            }
            let prg_width = ((prg_width_node !== undefined) ? ("width:" + (prg_width_node)) + "px;" : "width:inherit;");
            text += "<div style='height: 100%;direction: initial;overflow-wrap:break-word;word-wrap: break-word;" + prg_width + margin + "' >";
            text += prgrph_text;
            text += "</div>";
            text += "</div>";
        }

        return text;
    }

    private async genBuChar(node: any, i: number, spNode: any, textBodyNode: any, pFontStyle: any, idx: number | undefined, type: string | undefined): Promise<([string, number, number] | undefined)> {
        //console.log("genBuChar node: ", node, ", spNode: ", spNode, ", pFontStyle: ", pFontStyle, "type", type)
        ///////////////////////////////////////Amir///////////////////////////////
        let sldMstrTxtStyles = this.resources?.slideMasterTextStyles;
        let lstStyle = textBodyNode["a:lstStyle"];

        let rNode = getTextByPathList(node, ["a:r"]);
        if (rNode !== undefined && rNode.constructor === Array) {
            rNode = rNode[0]; //bullet only to first "a:r"
        }
        let lvl = parseInt(getTextByPathList(node["a:pPr"], ["attrs", "lvl"])) + 1;
        if (isNaN(lvl)) {
            lvl = 1;
        }
        let lvlStr = "a:lvl" + lvl + "pPr";
        let dfltBultColor, dfltBultSize, bultColor, bultSize, color_tye;

        if (rNode !== undefined) {
            dfltBultColor = this.getFontColorPr(rNode, spNode, lstStyle, pFontStyle, lvl, idx, type);
            color_tye = dfltBultColor[2];
            dfltBultSize = this.getFontSize(rNode, textBodyNode, pFontStyle, lvl, type);
        } else {
            return;
        }
        //console.log("Bullet Size: " + bultSize);

        let bullet = "", marRStr = "", marLStr = "", margin_val = 0, font_val = 0;
        /////////////////////////////////////////////////////////////////


        let pPrNode = node["a:pPr"];
        let BullNONE = getTextByPathList(pPrNode, ["a:buNone"]);
        if (BullNONE !== undefined) {
            return;
        }

        let buType = "TYPE_NONE";

        let layoutMasterNode = this.getLayoutAndMasterNode(node, idx, type);
        let pPrNodeLaout = layoutMasterNode.nodeLaout;
        let pPrNodeMaster = layoutMasterNode.nodeMaster;

        let buChar = getTextByPathList(pPrNode, ["a:buChar", "attrs", "char"]);
        let buNum = getTextByPathList(pPrNode, ["a:buAutoNum", "attrs", "type"]);
        let buPic = getTextByPathList(pPrNode, ["a:buBlip"]);
        if (buChar !== undefined) {
            buType = "TYPE_BULLET";
        }
        if (buNum !== undefined) {
            buType = "TYPE_NUMERIC";
        }
        if (buPic !== undefined) {
            buType = "TYPE_BULPIC";
        }

        let buFontSize = getTextByPathList(pPrNode, ["a:buSzPts", "attrs", "val"]);
        if (buFontSize === undefined) {
            buFontSize = getTextByPathList(pPrNode, ["a:buSzPct", "attrs", "val"]);
            if (buFontSize !== undefined) {
                let prcnt = parseInt(buFontSize) / 100000;
                //dfltBultSize = XXpt
                //var dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                let dfltBultSizeNoPt = parseInt(dfltBultSize);
                bultSize = prcnt * dfltBultSizeNoPt + "px";// + "pt";
            }
        } else {
            bultSize = (parseInt(buFontSize) / 100) * this.options.fontSizeFactor + "px";
        }

        //get definde bullet COLOR
        let buClrNode = getTextByPathList(pPrNode, ["a:buClr"]);


        if (buChar === undefined && buNum === undefined && buPic === undefined) {

            if (lstStyle !== undefined) {
                BullNONE = getTextByPathList(lstStyle, [lvlStr, "a:buNone"]);
                if (BullNONE !== undefined) {
                    return;
                }
                buType = "TYPE_NONE";
                buChar = getTextByPathList(lstStyle, [lvlStr, "a:buChar", "attrs", "char"]);
                buNum = getTextByPathList(lstStyle, [lvlStr, "a:buAutoNum", "attrs", "type"]);
                buPic = getTextByPathList(lstStyle, [lvlStr, "a:buBlip"]);
                if (buChar !== undefined) {
                    buType = "TYPE_BULLET";
                }
                if (buNum !== undefined) {
                    buType = "TYPE_NUMERIC";
                }
                if (buPic !== undefined) {
                    buType = "TYPE_BULPIC";
                }
                if (buChar !== undefined || buNum !== undefined || buPic !== undefined) {
                    pPrNode = lstStyle[lvlStr];
                }
            }
        }
        if (buChar === undefined && buNum === undefined && buPic === undefined) {
            //check in slidelayout and masterlayout - TODO
            if (pPrNodeLaout !== undefined) {
                BullNONE = getTextByPathList(pPrNodeLaout, ["a:buNone"]);
                if (BullNONE !== undefined) {
                    return;
                }
                buType = "TYPE_NONE";
                buChar = getTextByPathList(pPrNodeLaout, ["a:buChar", "attrs", "char"]);
                buNum = getTextByPathList(pPrNodeLaout, ["a:buAutoNum", "attrs", "type"]);
                buPic = getTextByPathList(pPrNodeLaout, ["a:buBlip"]);
                if (buChar !== undefined) {
                    buType = "TYPE_BULLET";
                }
                if (buNum !== undefined) {
                    buType = "TYPE_NUMERIC";
                }
                if (buPic !== undefined) {
                    buType = "TYPE_BULPIC";
                }
            }
            if (buChar === undefined && buNum === undefined && buPic === undefined) {
                //masterlayout

                if (pPrNodeMaster !== undefined) {
                    BullNONE = getTextByPathList(pPrNodeMaster, ["a:buNone"]);
                    if (BullNONE !== undefined) {
                        return;
                    }
                    buType = "TYPE_NONE";
                    buChar = getTextByPathList(pPrNodeMaster, ["a:buChar", "attrs", "char"]);
                    buNum = getTextByPathList(pPrNodeMaster, ["a:buAutoNum", "attrs", "type"]);
                    buPic = getTextByPathList(pPrNodeMaster, ["a:buBlip"]);
                    if (buChar !== undefined) {
                        buType = "TYPE_BULLET";
                    }
                    if (buNum !== undefined) {
                        buType = "TYPE_NUMERIC";
                    }
                    if (buPic !== undefined) {
                        buType = "TYPE_BULPIC";
                    }
                }

            }

        }
        //rtl
        let getRtlVal = getTextByPathList(pPrNode, ["attrs", "rtl"]);
        if (getRtlVal === undefined) {
            getRtlVal = getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
            if (getRtlVal === undefined && type != "shape") {
                getRtlVal = getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
            }
        }
        let isRTL = false;
        if (getRtlVal !== undefined && getRtlVal == "1") {
            isRTL = true;
        }
        //align
        let alignNode = getTextByPathList(pPrNode, ["attrs", "algn"]); //"l" | "ctr" | "r" | "just" | "justLow" | "dist" | "thaiDist
        if (alignNode === undefined) {
            alignNode = getTextByPathList(pPrNodeLaout, ["attrs", "algn"]);
            if (alignNode === undefined) {
                alignNode = getTextByPathList(pPrNodeMaster, ["attrs", "algn"]);
            }
        }
        //indent?
        let indentNode = getTextByPathList(pPrNode, ["attrs", "indent"]);
        if (indentNode === undefined) {
            indentNode = getTextByPathList(pPrNodeLaout, ["attrs", "indent"]);
            if (indentNode === undefined) {
                indentNode = getTextByPathList(pPrNodeMaster, ["attrs", "indent"]);
            }
        }
        let indent = 0;
        if (indentNode !== undefined) {
            indent = parseInt(indentNode) * this.options.slideFactor;
        }
        //marL
        let marLNode = getTextByPathList(pPrNode, ["attrs", "marL"]);
        if (marLNode === undefined) {
            marLNode = getTextByPathList(pPrNodeLaout, ["attrs", "marL"]);
            if (marLNode === undefined) {
                marLNode = getTextByPathList(pPrNodeMaster, ["attrs", "marL"]);
            }
        }
        //console.log("genBuChar() isRTL", isRTL, "alignNode:", alignNode)
        if (marLNode !== undefined) {
            let marginLeft = parseInt(marLNode) * this.options.slideFactor;
            if (isRTL) {// && alignNode == "r") {
                marLStr = "padding-right:";// "margin-right: ";
            } else {
                marLStr = "padding-left:";//"margin-left: ";
            }
            margin_val = ((marginLeft + indent < 0) ? 0 : (marginLeft + indent));
            marLStr += margin_val + "px;";
        }

        //marR?
        let marRNode = getTextByPathList(pPrNode, ["attrs", "marR"]);
        if (marRNode === undefined && marLNode === undefined) {
            //need to check if this posble - TODO
            marRNode = getTextByPathList(pPrNodeLaout, ["attrs", "marR"]);
            if (marRNode === undefined) {
                marRNode = getTextByPathList(pPrNodeMaster, ["attrs", "marR"]);
            }
        }
        if (marRNode !== undefined) {
            let marginRight = parseInt(marRNode) * this.options.slideFactor;
            if (isRTL) {// && alignNode == "r") {
                marLStr = "padding-right:";// "margin-right: ";
            } else {
                marLStr = "padding-left:";//"margin-left: ";
            }
            marRStr += ((marginRight + indent < 0) ? 0 : (marginRight + indent)) + "px;";
        }

        if (buType != "TYPE_NONE") {
            //var buFontAttrs = getTextByPathList(pPrNode, ["a:buFont", "attrs"]);
        }
        //console.log("Bullet Type: " + buType);
        //console.log("NumericTypr: " + buNum);
        //console.log("buChar: " + (buChar === undefined?'':buChar.charCodeAt(0)));
        //get definde bullet COLOR
        if (buClrNode === undefined) {
            //lstStyle
            buClrNode = getTextByPathList(lstStyle, [lvlStr, "a:buClr"]);
        }
        if (buClrNode === undefined) {
            buClrNode = getTextByPathList(pPrNodeLaout, ["a:buClr"]);
            if (buClrNode === undefined) {
                buClrNode = getTextByPathList(pPrNodeMaster, ["a:buClr"]);
            }
        }
        let defBultColor;
        if (buClrNode !== undefined) {
            defBultColor = this.getSolidFill(buClrNode);
        } else {
            if (pFontStyle !== undefined) {
                //console.log("genBuChar pFontStyle: ", pFontStyle)
                defBultColor = this.getSolidFill(pFontStyle);
            }
        }
        if (defBultColor === undefined || defBultColor == "NONE") {
            bultColor = dfltBultColor;
        } else {
            bultColor = [defBultColor, "", "solid"];
            color_tye = "solid";
        }
        //console.log("genBuChar node:", node, "pPrNode", pPrNode, " buClrNode: ", buClrNode, "defBultColor:", defBultColor,"dfltBultColor:" , dfltBultColor , "bultColor:", bultColor)

        //console.log("genBuChar: buClrNode: ", buClrNode, "bultColor", bultColor)
        //get definde bullet SIZE
        if (buFontSize === undefined) {
            buFontSize = getTextByPathList(pPrNodeLaout, ["a:buSzPts", "attrs", "val"]);
            if (buFontSize === undefined) {
                buFontSize = getTextByPathList(pPrNodeLaout, ["a:buSzPct", "attrs", "val"]);
                if (buFontSize !== undefined) {
                    let prcnt = parseInt(buFontSize) / 100000;
                    //var dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                    let dfltBultSizeNoPt = parseInt(dfltBultSize);
                    bultSize = prcnt * dfltBultSizeNoPt + "px";// + "pt";
                }
            } else {
                bultSize = (parseInt(buFontSize) / 100) * this.options.fontSizeFactor + "px";
            }
        }
        if (buFontSize === undefined) {
            buFontSize = getTextByPathList(pPrNodeMaster, ["a:buSzPts", "attrs", "val"]);
            if (buFontSize === undefined) {
                buFontSize = getTextByPathList(pPrNodeMaster, ["a:buSzPct", "attrs", "val"]);
                if (buFontSize !== undefined) {
                    let prcnt = parseInt(buFontSize) / 100000;
                    //dfltBultSize = XXpt
                    //var dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                    let dfltBultSizeNoPt = parseInt(dfltBultSize);
                    bultSize = prcnt * dfltBultSizeNoPt + "px";// + "pt";
                }
            } else {
                bultSize = (parseInt(buFontSize) / 100) * this.options.fontSizeFactor + "px";
            }
        }
        if (buFontSize === undefined) {
            bultSize = dfltBultSize;
        }
        font_val = parseInt(bultSize!);
        ////////////////////////////////////////////////////////////////////////
        if (buType == "TYPE_BULLET") {
            let typefaceNode = getTextByPathList(pPrNode, ["a:buFont", "attrs", "typeface"]);
            let typeface = "";
            if (typefaceNode !== undefined) {
                typeface = "font-family: " + typefaceNode;
            }
            // let marginLeft = parseInt(getTextByPathList(marLNode)) * this.options.slideFactor;
            // let marginRight = parseInt(getTextByPathList(marRNode)) * this.options.slideFactor;
            // if (isNaN(marginLeft)) {
            //     marginLeft = 328600 * this.options.slideFactor;
            // }
            // if (isNaN(marginRight)) {
            //     marginRight = 0;
            // }

            bullet = "<div style='height: 100%;" + typeface + ";" +
                marLStr + marRStr +
                "font-size:" + bultSize + ";";

            //bullet += "display: table-cell;";
            //"line-height: 0px;";
            if (color_tye == "solid") {
                if (bultColor[0] !== undefined && bultColor[0] != "") {
                    bullet += "color:#" + bultColor[0] + "; ";
                }
                if (bultColor[1] !== undefined && bultColor[1] != "" && bultColor[1] != ";") {
                    bullet += "text-shadow:" + bultColor[1] + ";";
                }
                //no highlight/background-color to bullet
                // if (bultColor[3] !== undefined && bultColor[3] != "") {
                //     styleText += "background-color: #" + bultColor[3] + ";";
                // }
            } else if (color_tye == "pattern" || color_tye == "pic" || color_tye == "gradient") {
                if (color_tye == "pattern") {
                    bullet += "background:" + bultColor[0][0] + ";";
                    if (bultColor[0][1] !== null && bultColor[0][1] !== undefined && bultColor[0][1] != "") {
                        bullet += "background-size:" + bultColor[0][1] + ";";//" 2px 2px;" +
                    }
                    if (bultColor[0][2] !== null && bultColor[0][2] !== undefined && bultColor[0][2] != "") {
                        bullet += "background-position:" + bultColor[0][2] + ";";//" 2px 2px;" +
                    }
                    // bullet += "-webkit-background-clip: text;" +
                    //     "background-clip: text;" +
                    //     "color: transparent;" +
                    //     "-webkit-text-stroke: " + bultColor[1].border + ";" +
                    //     "filter: " + bultColor[1].effcts + ";";
                } else if (color_tye == "pic") {
                    bullet += bultColor[0] + ";";
                    // bullet += "-webkit-background-clip: text;" +
                    //     "background-clip: text;" +
                    //     "color: transparent;" +
                    //     "-webkit-text-stroke: " + bultColor[1].border + ";";

                } else if (color_tye == "gradient") {

                    let colorAry = bultColor[0].color;
                    let rot = bultColor[0].rot;

                    bullet += "background: linear-gradient(" + rot + "deg,";
                    for (var i = 0; i < colorAry.length; i++) {
                        if (i == colorAry.length - 1) {
                            bullet += "#" + colorAry[i] + ");";
                        } else {
                            bullet += "#" + colorAry[i] + ", ";
                        }
                    }
                    // bullet += "color: transparent;" +
                    //     "-webkit-background-clip: text;" +
                    //     "background-clip: text;" +
                    //     "-webkit-text-stroke: " + bultColor[1].border + ";";
                }
                bullet += "-webkit-background-clip: text;" +
                    "background-clip: text;" +
                    "color: transparent;";
                if (bultColor[1].border !== undefined && bultColor[1].border !== "") {
                    bullet += "-webkit-text-stroke: " + bultColor[1].border + ";";
                }
                if (bultColor[1].effcts !== undefined && bultColor[1].effcts !== "") {
                    bullet += "filter: " + bultColor[1].effcts + ";";
                }
            }

            if (isRTL) {
                //bullet += "display: inline-block;white-space: nowrap ;direction:rtl"; // float: right;  
                bullet += "white-space: nowrap ;direction:rtl"; // display: table-cell;;
            }
            // @ts-ignore
            // TODO: supportIE 11 ?
            let isIE11 = !!window.MSInputMethodContext && !!document.documentMode;
            let htmlBu = buChar;

            if (!isIE11) {
                //ie11 does not support unicode ?
                htmlBu = getHtmlBullet(typefaceNode, buChar);
            }
            bullet += "'><div style='line-height: " + (font_val / 2) + "px;'>" + htmlBu + "</div></div>"; //font_val
            //} 
            // else {
            //     marginLeft = 328600 * this.options.slideFactor * lvl;

            //     bullet = "<div style='" + marLStr + "'>" + buChar + "</div>";
            // }
        } else if (buType == "TYPE_NUMERIC") { ///////////Amir///////////////////////////////
            //if (buFontAttrs !== undefined) {
            // let marginLeft = parseInt(getTextByPathList(pPrNode, ["attrs", "marL"])) * this.options.slideFactor;
            // let marginRight = parseInt(buFontAttrs["pitchFamily"]);

            // if (isNaN(marginLeft)) {
            //     marginLeft = 328600 * this.options.slideFactor;
            // }
            // if (isNaN(marginRight)) {
            //     marginRight = 0;
            // }
            //var typeface = buFontAttrs["typeface"];

            bullet = "<div style='height: 100%;" + marLStr + marRStr +
                "color:#" + bultColor[0] + ";" +
                "font-size:" + bultSize + ";";// +
            //"line-height: 0px;";
            if (isRTL) {
                bullet += "display: inline-block;white-space: nowrap ;direction:rtl;"; // float: right;
            } else {
                bullet += "display: inline-block;white-space: nowrap ;direction:ltr;"; //float: left;
            }
            bullet += "' data-bulltname = '" + buNum + "' data-bulltlvl = '" + lvl + "' class='numeric-bullet-style'></div>";
            // } else {
            //     marginLeft = 328600 * this.options.slideFactor * lvl;
            //     bullet = "<div style='margin-left: " + marginLeft + "px;";
            //     if (isRTL) {
            //         bullet += " float: right; direction:rtl;";
            //     } else {
            //         bullet += " float: left; direction:ltr;";
            //     }
            //     bullet += "' data-bulltname = '" + buNum + "' data-bulltlvl = '" + lvl + "' class='numeric-bullet-style'></div>";
            // }

        } else if (buType == "TYPE_BULPIC") { //PIC BULLET
            // let marginLeft = parseInt(getTextByPathList(pPrNode, ["attrs", "marL"])) * this.options.slideFactor;
            // let marginRight = parseInt(getTextByPathList(pPrNode, ["attrs", "marR"])) * this.options.slideFactor;

            // if (isNaN(marginRight)) {
            //     marginRight = 0;
            // }
            // //console.log("marginRight: "+marginRight)
            // //buPic
            // if (isNaN(marginLeft)) {
            //     marginLeft = 328600 * this.options.slideFactor;
            // } else {
            //     marginLeft = 0;
            // }
            //var buPicId = getTextByPathList(buPic, ["a:blip","a:extLst","a:ext","asvg:svgBlip" , "attrs", "r:embed"]);
            let buPicId = getTextByPathList(buPic, ["a:blip", "attrs", "r:embed"]);
            let svgPicPath = "";
            let buImg;
            if (buPicId !== undefined) {
                //svgPicPath = this.resources["slideResObj"][buPicId]["target"];
                //buImg = this.resources["zip"].file(svgPicPath).asText();
                //}else{
                //buPicId = getTextByPathList(buPic, ["a:blip", "attrs", "r:embed"]);
                let imgPath = this.resources?.slideResObj[buPicId]["target"];
                //console.log("imgPath: ", imgPath);
                let imgBase64 = await this.options.resourceReader?.getImgBase64Data(imgPath);
                if (imgBase64) {
                    buImg = "<img src='" + imgBase64 + "' style='width: 100%;'/>"// height: 100%
                    //console.log("imgPath: "+imgPath+"\nimgMimeType: "+imgMimeType)
                }
            }
            if (buPicId === undefined) {
                buImg = "&#8227;";
            }
            bullet = "<div style='height: 100%;" + marLStr + marRStr +
                "width:" + bultSize + ";display: inline-block; ";// +
            //"line-height: 0px;";
            if (isRTL) {
                bullet += "display: inline-block;white-space: nowrap ;direction:rtl;"; //direction:rtl; float: right;
            }
            bullet += "'>" + buImg + "  </div>";
            //////////////////////////////////////////////////////////////////////////////////////
        }
        // else {
        //     bullet = "<div style='margin-left: " + 328600 * this.options.slideFactor * lvl + "px" +
        //         "; margin-right: " + 0 + "px;'></div>";
        // }
        //console.log("genBuChar: width: ", $(bullet).outerWidth())
        return [bullet, margin_val, font_val];//$(bullet).outerWidth()];
    }

    private getLayoutAndMasterNode(node: any, idx: number | undefined, type: string | undefined) {
        let pPrNodeLaout, pPrNodeMaster;
        let pPrNode = node["a:pPr"];
        //lvl
        let lvl = 1;
        let lvlNode = getTextByPathList(pPrNode, ["attrs", "lvl"]);
        if (lvlNode !== undefined) {
            lvl = parseInt(lvlNode) + 1;
        }
        const slideLayoutTables = this.resources?.slideLayoutTables;
        const slideMasterTextStyles = this.resources?.slideMasterTextStyles;

        if (idx !== undefined) {
            //slidelayout
            pPrNodeLaout = getTextByPathList(slideLayoutTables["idxTable"][idx], ["p:txBody", "a:lstStyle", "a:lvl" + lvl + "pPr"]);
            if (pPrNodeLaout === undefined) {
                pPrNodeLaout = getTextByPathList(slideLayoutTables["idxTable"][idx], ["p:txBody", "a:p", "a:pPr"]);
                if (pPrNodeLaout === undefined) {
                    pPrNodeLaout = getTextByPathList(slideLayoutTables["idxTable"][idx], ["p:txBody", "a:p", (lvl - 1), "a:pPr"]);
                }
            }
        }
        if (type !== undefined) {
            //slidelayout
            let lvlStr = "a:lvl" + lvl + "pPr";
            if (pPrNodeLaout === undefined) {
                pPrNodeLaout = getTextByPathList(slideLayoutTables, ["typeTable", type, "p:txBody", "a:lstStyle", lvlStr]);
            }
            //masterlayout
            if (type == "title" || type == "ctrTitle") {
                pPrNodeMaster = getTextByPathList(slideMasterTextStyles, ["p:titleStyle", lvlStr]);
            } else if (type == "body" || type == "obj" || type == "subTitle") {
                pPrNodeMaster = getTextByPathList(slideMasterTextStyles, ["p:bodyStyle", lvlStr]);
            } else if (type == "shape" || type == "diagram") {
                pPrNodeMaster = getTextByPathList(slideMasterTextStyles, ["p:otherStyle", lvlStr]);
            } else if (type == "textBox") {
                pPrNodeMaster = getTextByPathList(this.options.basicInfo?.defaultTextStyle, [lvlStr]);
            } else {
                pPrNodeMaster = getTextByPathList(this.resources?.slideMasterTables, ["typeTable", type, "p:txBody", "a:lstStyle", lvlStr]);
            }
        }
        return {
            "nodeLaout": pPrNodeLaout,
            "nodeMaster": pPrNodeMaster
        };
    }
    private genSpanElement(node: any, rIndex: number | undefined, pNode: any, textBodyNode: any, pFontStyle: any, slideLayoutSpNode: any, idx: number | undefined, type: string | undefined, rNodeLength: number, isBullate: boolean) {
        //https://codepen.io/imdunn/pen/GRgwaye ?
        let text_style = "";
        let lstStyle = textBodyNode["a:lstStyle"];

        let text = node["a:t"];
        //var text_count = text.length;

        let openElemnt = "<span";//"<bdi";
        let closeElemnt = "</span>";// "</bdi>";
        let styleText = "";
        if (text === undefined && node["type"] !== undefined) {
            if (this.isFirstBr) {
                //openElemnt = "<br";
                //closeElemnt = "";
                //return "<br style='font-size: initial'>"
                this.isFirstBr = false;
                return "<span class='line-break-br' ></span>";
            } else {
                // styleText += "display: block;";
                // openElemnt = "<sapn";
                // closeElemnt = "</sapn>";
            }

            styleText += "display: block;";
            //openElemnt = "<sapn";
            //closeElemnt = "</sapn>";
        } else {
            this.isFirstBr = true;
        }
        if (typeof text !== 'string') {
            text = getTextByPathList(node, ["a:fld", "a:t"]);
            if (typeof text !== 'string') {
                text = "&nbsp;";
                //return "<span class='text-block '>&nbsp;</span>";
            }
            // if (text === undefined) {
            //     return "";
            // }
        }

        let pPrNode = pNode["a:pPr"];
        //lvl
        let lvl = 1;
        let lvlNode = getTextByPathList(pPrNode, ["attrs", "lvl"]);
        if (lvlNode !== undefined) {
            lvl = parseInt(lvlNode) + 1;
        }
        //console.log("genSpanElement node: ", node, "rIndex: ", rIndex, ", pNode: ", pNode, ",pPrNode: ", pPrNode, "pFontStyle:", pFontStyle, ", idx: ", idx, "type:", type, this.resources);
        let layoutMasterNode = this.getLayoutAndMasterNode(pNode, idx, type);
        let pPrNodeLaout = layoutMasterNode.nodeLaout;
        let pPrNodeMaster = layoutMasterNode.nodeMaster;

        //Language
        let lang = getTextByPathList(node, ["a:rPr", "attrs", "lang"]);
        let isRtlLan = (lang !== undefined && rtlLangs.indexOf(lang) !== -1) ? true : false;
        //rtl
        let getRtlVal = getTextByPathList(pPrNode, ["attrs", "rtl"]);
        if (getRtlVal === undefined) {
            getRtlVal = getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
            if (getRtlVal === undefined && type != "shape") {
                getRtlVal = getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
            }
        }
        let isRTL = false;
        let dirStr = "ltr";
        if (getRtlVal !== undefined && getRtlVal == "1") {
            isRTL = true;
            dirStr = "rtl";
        }

        let linkID = getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "r:id"]);
        let linkTooltip = "";
        let defLinkClr;
        if (linkID !== undefined) {
            linkTooltip = getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "tooltip"]);
            if (linkTooltip !== undefined) {
                linkTooltip = "title='" + linkTooltip + "'";
            }
            defLinkClr = this.getSchemeColorFromTheme("a:hlink");

            let linkClrNode = getTextByPathList(node, ["a:rPr", "a:solidFill"]);// getTextByPathList(node, ["a:rPr", "a:solidFill"]);
            let rPrlinkClr = this.getSolidFill(linkClrNode);


            //console.log("genSpanElement defLinkClr: ", defLinkClr, "rPrlinkClr:", rPrlinkClr)
            if (rPrlinkClr !== undefined && rPrlinkClr != "") {
                defLinkClr = rPrlinkClr;
            }

        }
        /////////////////////////////////////////////////////////////////////////////////////
        //getFontColor
        let fontClrPr = this.getFontColorPr(node, pNode, lstStyle, pFontStyle, lvl, idx, type);
        let fontClrType = fontClrPr[2];
        //console.log("genSpanElement fontClrPr: ", fontClrPr, "linkID", linkID);
        if (fontClrType == "solid") {
            if (linkID === undefined && fontClrPr[0] !== undefined && fontClrPr[0] != "") {
                styleText += "color: #" + fontClrPr[0] + ";";
            }
            else if (linkID !== undefined && defLinkClr !== undefined) {
                styleText += "color: #" + defLinkClr + ";";
            }

            if (fontClrPr[1] !== undefined && fontClrPr[1] != "" && fontClrPr[1] != ";") {
                styleText += "text-shadow:" + fontClrPr[1] + ";";
            }
            if (fontClrPr[3] !== undefined && fontClrPr[3] != "") {
                styleText += "background-color: #" + fontClrPr[3] + ";";
            }
        } else if (fontClrType == "pattern" || fontClrType == "pic" || fontClrType == "gradient") {
            if (fontClrType == "pattern") {
                styleText += "background:" + fontClrPr[0][0] + ";";
                if (fontClrPr[0][1] !== null && fontClrPr[0][1] !== undefined && fontClrPr[0][1] != "") {
                    styleText += "background-size:" + fontClrPr[0][1] + ";";//" 2px 2px;" +
                }
                if (fontClrPr[0][2] !== null && fontClrPr[0][2] !== undefined && fontClrPr[0][2] != "") {
                    styleText += "background-position:" + fontClrPr[0][2] + ";";//" 2px 2px;" +
                }
                // styleText += "-webkit-background-clip: text;" +
                //     "background-clip: text;" +
                //     "color: transparent;" +
                //     "-webkit-text-stroke: " + fontClrPr[1].border + ";" +
                //     "filter: " + fontClrPr[1].effcts + ";";
            } else if (fontClrType == "pic") {
                styleText += fontClrPr[0] + ";";
                // styleText += "-webkit-background-clip: text;" +
                //     "background-clip: text;" +
                //     "color: transparent;" +
                //     "-webkit-text-stroke: " + fontClrPr[1].border + ";";
            } else if (fontClrType == "gradient") {

                let colorAry = fontClrPr[0].color;
                let rot = fontClrPr[0].rot;

                styleText += "background: linear-gradient(" + rot + "deg,";
                for (var i = 0; i < colorAry.length; i++) {
                    if (i == colorAry.length - 1) {
                        styleText += "#" + colorAry[i] + ");";
                    } else {
                        styleText += "#" + colorAry[i] + ", ";
                    }
                }
                // styleText += "-webkit-background-clip: text;" +
                //     "background-clip: text;" +
                //     "color: transparent;" +
                //     "-webkit-text-stroke: " + fontClrPr[1].border + ";";

            }
            styleText += "-webkit-background-clip: text;" +
                "background-clip: text;" +
                "color: transparent;";
            if (fontClrPr[1].border !== undefined && fontClrPr[1].border !== "") {
                styleText += "-webkit-text-stroke: " + fontClrPr[1].border + ";";
            }
            if (fontClrPr[1].effcts !== undefined && fontClrPr[1].effcts !== "") {
                styleText += "filter: " + fontClrPr[1].effcts + ";";
            }
        }
        let font_size = this.getFontSize(node, textBodyNode, pFontStyle, lvl, type);
        //text_style += "font-size:" + font_size + ";"

        text_style += "font-size:" + font_size + ";" +
            // marLStr +
            "font-family:" + this.getFontType(node, type, pFontStyle) + ";" +
            "font-weight:" + this.getFontBold(node) + ";" +
            "font-style:" + this.getFontItalic(node) + ";" +
            "text-decoration:" + this.getFontDecoration(node) + ";" +
            "text-align:" + this.getTextHorizontalAlign(node, pNode, type) + ";" +
            "vertical-align:" + this.getTextVerticalAlign(node) + ";";
        //rNodeLength
        //console.log("genSpanElement node:", node, "lang:", lang, "isRtlLan:", isRtlLan, "span parent dir:", dirStr)
        if (isRtlLan) { //|| rIndex === undefined
            styleText += "direction:rtl;";
        } else { //|| rIndex === undefined
            styleText += "direction:ltr;";
        }
        // } else if (dirStr == "rtl" && isRtlLan ) {
        //     styleText += "direction:rtl;";

        // } else if (dirStr == "ltr" && !isRtlLan ) {
        //     styleText += "direction:ltr;";
        // } else if (dirStr == "ltr" && isRtlLan){
        //     styleText += "direction:ltr;";
        // }else{
        //     styleText += "direction:inherit;";
        // }

        // if (dirStr == "rtl" && !isRtlLan) { //|| rIndex === undefined
        //     styleText += "direction:ltr;";
        // } else if (dirStr == "rtl" && isRtlLan) {
        //     styleText += "direction:rtl;";
        // } else if (dirStr == "ltr" && !isRtlLan) {
        //     styleText += "direction:ltr;";
        // } else if (dirStr == "ltr" && isRtlLan) {
        //     styleText += "direction:rtl;";
        // } else {
        //     styleText += "direction:inherit;";
        // }

        //     //"direction:" + dirStr + ";";
        //if (rNodeLength == 1 || rIndex == 0 ){
        //styleText += "display: table-cell;white-space: nowrap;";
        //}
        let highlight = getTextByPathList(node, ["a:rPr", "a:highlight"]);
        if (highlight !== undefined) {
            styleText += "background-color:#" + this.getSolidFill(highlight) + ";";
            //styleText += "Opacity:" + getColorOpacity(highlight) + ";";
        }

        //letter-spacing:
        let spcNode = getTextByPathList(node, ["a:rPr", "attrs", "spc"]);
        if (spcNode === undefined) {
            spcNode = getTextByPathList(pPrNodeLaout, ["a:defRPr", "attrs", "spc"]);
            if (spcNode === undefined) {
                spcNode = getTextByPathList(pPrNodeMaster, ["a:defRPr", "attrs", "spc"]);
            }
        }
        if (spcNode !== undefined) {
            let ltrSpc = parseInt(spcNode) / 100; //pt
            styleText += "letter-spacing: " + ltrSpc + "px;";// + "pt;";
        }

        //Text Cap Types
        let capNode = getTextByPathList(node, ["a:rPr", "attrs", "cap"]);
        if (capNode === undefined) {
            capNode = getTextByPathList(pPrNodeLaout, ["a:defRPr", "attrs", "cap"]);
            if (capNode === undefined) {
                capNode = getTextByPathList(pPrNodeMaster, ["a:defRPr", "attrs", "cap"]);
            }
        }
        if (capNode == "small" || capNode == "all") {
            styleText += "text-transform: uppercase";
        }
        //styleText += "word-break: break-word;";
        //console.log("genSpanElement node: ", node, ", capNode: ", capNode, ",pPrNodeLaout: ", pPrNodeLaout, ", pPrNodeMaster: ", pPrNodeMaster, "this.resources:", this.resources);

        let cssName = "";

        if (this.options.styleTable?.has(styleText)) {
            cssName = this.options.styleTable.get(styleText).name;
        } else {
            cssName = "_css_" + this.options.styleTable?.nextIndex();
            this.options.styleTable?.update(styleText, {
                "name": cssName,
                "text": styleText
            });
        }
        let linkColorSyle = "";
        if (fontClrType == "solid" && linkID !== undefined) {
            linkColorSyle = "style='color: inherit;'";
        }

        if (linkID !== undefined && linkID != "") {
            let linkURL = this.resources?.["slideResObj"][linkID]["target"];
            linkURL = escapeHtml(linkURL);
            return openElemnt + " class='text-block " + cssName + "' style='" + text_style + "'><a href='" + linkURL + "' " + linkColorSyle + "  " + linkTooltip + " target='_blank'>" +
                escapeHtml(text) + "</a>" + closeElemnt;
        } else {
            // charactors like '<' should be escaped incase of it is a html tag
            return openElemnt + " class='text-block " + cssName + "' style='" + text_style + "'>" + escapeHtml(text) + closeElemnt;//"</bdi>";
        }

    }

    private getPregraphMargn(pNode: any, idx: number | undefined, type: string | undefined, isBullate: boolean) {
        if (!isBullate) {
            return ["", 0];
        }
        let marLStr = "", marRStr = "", maginVal = 0;
        let pPrNode = pNode["a:pPr"];
        let layoutMasterNode = this.getLayoutAndMasterNode(pNode, idx, type);
        let pPrNodeLaout = layoutMasterNode.nodeLaout;
        let pPrNodeMaster = layoutMasterNode.nodeMaster;

        // let lang = getTextByPathList(node, ["a:rPr", "attrs", "lang"]);
        // let isRtlLan = (lang !== undefined && rtl_langs_array.indexOf(lang) !== -1) ? true : false;
        //rtl
        let getRtlVal = getTextByPathList(pPrNode, ["attrs", "rtl"]);
        if (getRtlVal === undefined) {
            getRtlVal = getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
            if (getRtlVal === undefined && type != "shape") {
                getRtlVal = getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
            }
        }
        let isRTL = false;
        let dirStr = "ltr";
        if (getRtlVal !== undefined && getRtlVal == "1") {
            isRTL = true;
            dirStr = "rtl";
        }

        //align
        let alignNode = getTextByPathList(pPrNode, ["attrs", "algn"]); //"l" | "ctr" | "r" | "just" | "justLow" | "dist" | "thaiDist
        if (alignNode === undefined) {
            alignNode = getTextByPathList(pPrNodeLaout, ["attrs", "algn"]);
            if (alignNode === undefined) {
                alignNode = getTextByPathList(pPrNodeMaster, ["attrs", "algn"]);
            }
        }
        //indent?
        let indentNode = getTextByPathList(pPrNode, ["attrs", "indent"]);
        if (indentNode === undefined) {
            indentNode = getTextByPathList(pPrNodeLaout, ["attrs", "indent"]);
            if (indentNode === undefined) {
                indentNode = getTextByPathList(pPrNodeMaster, ["attrs", "indent"]);
            }
        }
        let indent = 0;
        if (indentNode !== undefined) {
            indent = parseInt(indentNode) * this.options.slideFactor;
        }
        //
        //marL
        let marLNode = getTextByPathList(pPrNode, ["attrs", "marL"]);
        if (marLNode === undefined) {
            marLNode = getTextByPathList(pPrNodeLaout, ["attrs", "marL"]);
            if (marLNode === undefined) {
                marLNode = getTextByPathList(pPrNodeMaster, ["attrs", "marL"]);
            }
        }
        let marginLeft = 0;
        if (marLNode !== undefined) {
            marginLeft = parseInt(marLNode) * this.options.slideFactor;
        }
        if ((indentNode !== undefined || marLNode !== undefined)) {
            //var lvlIndent = defTabSz * lvl;

            if (isRTL) {// && alignNode == "r") {
                //marLStr = "margin-right: ";
                marLStr = "padding-right: ";
            } else {
                //marLStr = "margin-left: ";
                marLStr = "padding-left: ";
            }
            if (isBullate) {
                maginVal = Math.abs(0 - indent);
                marLStr += maginVal + "px;";  // (minus bullate numeric lenght/size - TODO
            } else {
                maginVal = Math.abs(marginLeft + indent);
                marLStr += maginVal + "px;";  // (minus bullate numeric lenght/size - TODO
            }
        }

        //marR?
        let marRNode = getTextByPathList(pPrNode, ["attrs", "marR"]);
        if (marRNode === undefined && marLNode === undefined) {
            //need to check if this posble - TODO
            marRNode = getTextByPathList(pPrNodeLaout, ["attrs", "marR"]);
            if (marRNode === undefined) {
                marRNode = getTextByPathList(pPrNodeMaster, ["attrs", "marR"]);
            }
        }
        if (marRNode !== undefined && isBullate) {
            let marginRight = parseInt(marRNode) * this.options.slideFactor;
            if (isRTL) {// && alignNode == "r") {
                //marRStr = "margin-right: ";
                marRStr = "padding-right: ";
            } else {
                //marRStr = "margin-left: ";
                marRStr = "padding-left: ";
            }
            marRStr += Math.abs(0 - indent) + "px;";
        }


        return [marLStr, maginVal];
    }

    private async genTable(node: any) {
        let order = node["attrs"]?.["order"];
        let tableNode = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl"]);
        let xfrmNode = getTextByPathList(node, ["p:xfrm"]);
        /////////////////////////////////////////Amir////////////////////////////////////////////////
        let getTblPr = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblPr"]);
        let getColsGrid = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblGrid", "a:gridCol"]);
        let tblDir = "";
        if (getTblPr !== undefined) {
            let isRTL = getTblPr["attrs"]?.["rtl"];
            tblDir = (isRTL == 1 ? "dir=rtl" : "dir=ltr");
        }
        let firstRowAttr = getTblPr["attrs"]?.["firstRow"]; //associated element <a:firstRow> in the table styles
        let firstColAttr = getTblPr["attrs"]?.["firstCol"]; //associated element <a:firstCol> in the table styles
        let lastRowAttr = getTblPr["attrs"]?.["lastRow"]; //associated element <a:lastRow> in the table styles
        let lastColAttr = getTblPr["attrs"]?.["lastCol"]; //associated element <a:lastCol> in the table styles
        let bandRowAttr = getTblPr["attrs"]?.["bandRow"]; //associated element <a:band1H>, <a:band2H> in the table styles
        let bandColAttr = getTblPr["attrs"]?.["bandCol"]; //associated element <a:band1V>, <a:band2V> in the table styles
        //console.log("getTblPr: ", getTblPr);
        let tblStylAttrObj = {
            isFrstRowAttr: (firstRowAttr !== undefined && firstRowAttr == "1") ? 1 : 0,
            isFrstColAttr: (firstColAttr !== undefined && firstColAttr == "1") ? 1 : 0,
            isLstRowAttr: (lastRowAttr !== undefined && lastRowAttr == "1") ? 1 : 0,
            isLstColAttr: (lastColAttr !== undefined && lastColAttr == "1") ? 1 : 0,
            isBandRowAttr: (bandRowAttr !== undefined && bandRowAttr == "1") ? 1 : 0,
            isBandColAttr: (bandColAttr !== undefined && bandColAttr == "1") ? 1 : 0
        }

        let thisTblStyle: any;
        let tbleStyleId = getTblPr["a:tableStyleId"];
        if (tbleStyleId !== undefined) {
            let tbleStylList = this.options.tableStyles?.["a:tblStyleLst"]?.["a:tblStyle"];
            if (tbleStylList !== undefined) {
                if (Array.isArray(tbleStylList)) {
                    for (var k = 0; k < (tbleStylList as any[]).length; k++) {
                        if (tbleStylList[k]["attrs"]?.["styleId"] == tbleStyleId) {
                            thisTblStyle = tbleStylList[k];
                        }
                    }
                } else {
                    if (tbleStylList["attrs"]?.["styleId"] == tbleStyleId) {
                        thisTblStyle = tbleStylList;
                    }
                }
            }
        }
        if (thisTblStyle !== undefined) {
            thisTblStyle["tblStylAttrObj"] = tblStylAttrObj;
            if (this.resources) {
                this.resources.tableStyles = thisTblStyle;
            }
        }
        let tblStyl = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle"]);
        let tblBorderStyl = getTextByPathList(tblStyl, ["a:tcBdr"]);
        let tbl_borders = "";
        if (tblBorderStyl !== undefined) {
            tbl_borders = this.getTableBorders(tblBorderStyl);
        }
        let tbl_bgcolor: string | undefined = "";
        let tbl_opacity = 1;
        let tbl_bgFillschemeClr = getTextByPathList(thisTblStyle, ["a:tblBg", "a:fillRef"]);
        //console.log( "thisTblStyle:", thisTblStyle, "this.resources:", this.resources)
        if (tbl_bgFillschemeClr !== undefined) {
            tbl_bgcolor = this.getSolidFill(tbl_bgFillschemeClr);
        }
        if (tbl_bgFillschemeClr === undefined) {
            tbl_bgFillschemeClr = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
            tbl_bgcolor = this.getSolidFill(tbl_bgFillschemeClr);
        }
        if (tbl_bgcolor !== "") {
            tbl_bgcolor = "background-color: #" + tbl_bgcolor + ";";
        }
        ////////////////////////////////////////////////////////////////////////////////////////////
        let tableHtml = "<table " + tblDir + " style='border-collapse: collapse;" +
            this.getPosition(xfrmNode, node, undefined, undefined) +
            this.getSize(xfrmNode, node, undefined, undefined) +
            " z-index: " + order + ";" +
            tbl_borders + ";" +
            tbl_bgcolor + "'>";

        let trNodes = tableNode["a:tr"];
        if (trNodes.constructor !== Array) {
            trNodes = [trNodes];
        }
        //if (trNodes.constructor === Array) {
        //multi rows
        let totalrowSpan = 0;
        let rowSpanAry: any[] = [];
        for (var i = 0; i < trNodes.length; i++) {
            //////////////rows Style ////////////Amir
            let rowHeightParam = trNodes[i]["attrs"]?.["h"];
            let rowHeight = 0;
            let rowsStyl = "";
            if (rowHeightParam !== undefined) {
                rowHeight = parseInt(rowHeightParam) * this.options.slideFactor;
                rowsStyl += "height:" + rowHeight + "px;";
            }
            let fillColor: string | undefined = "";
            let row_borders: string | undefined = "";
            let fontClrPr = "";
            let fontWeight = "";
            let band_1H_fillColor;
            let band_2H_fillColor;

            if (thisTblStyle !== undefined && thisTblStyle["a:wholeTbl"] !== undefined) {
                let bgFillschemeClr = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    let local_fillColor = this.getSolidFill(bgFillschemeClr);
                    if (local_fillColor !== undefined) {
                        fillColor = local_fillColor;
                    }
                }
                let rowTxtStyl = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
                if (rowTxtStyl !== undefined) {
                    let local_fontColor = this.getSolidFill(rowTxtStyl);
                    if (local_fontColor !== undefined) {
                        fontClrPr = local_fontColor;
                    }

                    let local_fontWeight = ((getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                    if (local_fontWeight != "") {
                        fontWeight = local_fontWeight
                    }
                }
            }

            if (i == 0 && tblStylAttrObj["isFrstRowAttr"] == 1 && thisTblStyle !== undefined) {

                let bgFillschemeClr = getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    let local_fillColor = this.getSolidFill(bgFillschemeClr);
                    if (local_fillColor !== undefined) {
                        fillColor = local_fillColor;
                    }
                }
                let borderStyl = getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:tcBdr"]);
                if (borderStyl !== undefined) {
                    let local_row_borders = this.getTableBorders(borderStyl);
                    if (local_row_borders != "") {
                        row_borders = local_row_borders;
                    }
                }
                let rowTxtStyl = getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcTxStyle"]);
                if (rowTxtStyl !== undefined) {
                    let local_fontClrPr = this.getSolidFill(rowTxtStyl);
                    if (local_fontClrPr !== undefined) {
                        fontClrPr = local_fontClrPr;
                    }
                    let local_fontWeight = ((getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                    if (local_fontWeight !== "") {
                        fontWeight = local_fontWeight;
                    }
                }

            } else if (i > 0 && tblStylAttrObj["isBandRowAttr"] == 1 && thisTblStyle !== undefined) {
                fillColor = "";
                row_borders = undefined;
                if ((i % 2) == 0 && thisTblStyle["a:band2H"] !== undefined) {
                    // console.log("i: ", i, 'thisTblStyle["a:band2H"]:', thisTblStyle["a:band2H"])
                    //check if there is a row bg
                    let bgFillschemeClr = getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:fill", "a:solidFill"]);
                    if (bgFillschemeClr !== undefined) {
                        let local_fillColor = this.getSolidFill(bgFillschemeClr);
                        if (local_fillColor !== "") {
                            fillColor = local_fillColor;
                            band_2H_fillColor = local_fillColor;
                        }
                    }


                    let borderStyl = getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:tcBdr"]);
                    if (borderStyl !== undefined) {
                        let local_row_borders = this.getTableBorders(borderStyl);
                        if (local_row_borders != "") {
                            row_borders = local_row_borders;
                        }
                    }
                    let rowTxtStyl = getTextByPathList(thisTblStyle, ["a:band2H", "a:tcTxStyle"]);
                    if (rowTxtStyl !== undefined) {
                        let local_fontClrPr = this.getSolidFill(rowTxtStyl);
                        if (local_fontClrPr !== undefined) {
                            fontClrPr = local_fontClrPr;
                        }
                    }

                    let local_fontWeight = ((getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");

                    if (local_fontWeight !== "") {
                        fontWeight = local_fontWeight;
                    }
                }
                if ((i % 2) != 0 && thisTblStyle["a:band1H"] !== undefined) {
                    let bgFillschemeClr = getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:fill", "a:solidFill"]);
                    if (bgFillschemeClr !== undefined) {
                        let local_fillColor = this.getSolidFill(bgFillschemeClr);
                        if (local_fillColor !== undefined) {
                            fillColor = local_fillColor;
                            band_1H_fillColor = local_fillColor;
                        }
                    }
                    let borderStyl = getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:tcBdr"]);
                    if (borderStyl !== undefined) {
                        let local_row_borders = this.getTableBorders(borderStyl);
                        if (local_row_borders != "") {
                            row_borders = local_row_borders;
                        }
                    }
                    let rowTxtStyl = getTextByPathList(thisTblStyle, ["a:band1H", "a:tcTxStyle"]);
                    if (rowTxtStyl !== undefined) {
                        let local_fontClrPr = this.getSolidFill(rowTxtStyl);
                        if (local_fontClrPr !== undefined) {
                            fontClrPr = local_fontClrPr;
                        }
                        let local_fontWeight = ((getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                        if (local_fontWeight != "") {
                            fontWeight = local_fontWeight;
                        }
                    }
                }

            }
            //last row
            if (i == (trNodes.length - 1) && tblStylAttrObj["isLstRowAttr"] == 1 && thisTblStyle !== undefined) {
                let bgFillschemeClr = getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    let local_fillColor = this.getSolidFill(bgFillschemeClr);
                    if (local_fillColor !== undefined) {
                        fillColor = local_fillColor;
                    }
                    // let local_colorOpacity = getColorOpacity(bgFillschemeClr);
                    // if(local_colorOpacity !== undefined){
                    //     colorOpacity = local_colorOpacity;
                    // }
                }
                let borderStyl = getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:tcBdr"]);
                if (borderStyl !== undefined) {
                    let local_row_borders = this.getTableBorders(borderStyl);
                    if (local_row_borders != "") {
                        row_borders = local_row_borders;
                    }
                }
                let rowTxtStyl = getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcTxStyle"]);
                if (rowTxtStyl !== undefined) {
                    let local_fontClrPr = this.getSolidFill(rowTxtStyl);
                    if (local_fontClrPr !== undefined) {
                        fontClrPr = local_fontClrPr;
                    }

                    let local_fontWeight = ((getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                    if (local_fontWeight !== "") {
                        fontWeight = local_fontWeight;
                    }
                }
            }
            rowsStyl += ((row_borders !== undefined) ? row_borders : "");
            rowsStyl += ((fontClrPr !== undefined) ? " color: #" + fontClrPr + ";" : "");
            rowsStyl += ((fontWeight != "") ? " font-weight:" + fontWeight + ";" : "");
            if (fillColor !== undefined && fillColor != "") {
                //rowsStyl += "background-color: rgba(" + hexToRgbNew(fillColor) + "," + colorOpacity + ");";
                rowsStyl += "background-color: #" + fillColor + ";";
            }
            tableHtml += "<tr style='" + rowsStyl + "'>";
            ////////////////////////////////////////////////

            let tcNodes = trNodes[i]["a:tc"];
            if (tcNodes !== undefined) {
                if (tcNodes.constructor === Array) {
                    //multi columns
                    let j = 0;
                    if (rowSpanAry.length == 0) {
                        rowSpanAry = Array.apply(null, Array(tcNodes.length)).map(function () { return 0 });
                    }
                    let totalColSpan = 0;
                    while (j < tcNodes.length) {
                        if (rowSpanAry[j] == 0 && totalColSpan == 0) {
                            let a_sorce;
                            //j=0 : first col
                            if (j == 0 && tblStylAttrObj["isFrstColAttr"] == 1) {
                                a_sorce = "a:firstCol";
                                if (tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1) &&
                                    getTextByPathList(thisTblStyle, ["a:seCell"]) !== undefined) {
                                    a_sorce = "a:seCell";
                                } else if (tblStylAttrObj["isFrstRowAttr"] == 1 && i == 0 &&
                                    getTextByPathList(thisTblStyle, ["a:neCell"]) !== undefined) {
                                    a_sorce = "a:neCell";
                                }
                            } else if ((j > 0 && tblStylAttrObj["isBandColAttr"] == 1) &&
                                !(tblStylAttrObj["isFrstColAttr"] == 1 && i == 0) &&
                                !(tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1)) &&
                                j != (tcNodes.length - 1)) {

                                if ((j % 2) != 0) {

                                    let aBandNode = getTextByPathList(thisTblStyle, ["a:band2V"]);
                                    if (aBandNode === undefined) {
                                        aBandNode = getTextByPathList(thisTblStyle, ["a:band1V"]);
                                        if (aBandNode !== undefined) {
                                            a_sorce = "a:band2V";
                                        }
                                    } else {
                                        a_sorce = "a:band2V";
                                    }

                                }
                            }

                            if (j == (tcNodes.length - 1) && tblStylAttrObj["isLstColAttr"] == 1) {
                                a_sorce = "a:lastCol";
                                if (tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1) && getTextByPathList(thisTblStyle, ["a:swCell"]) !== undefined) {
                                    a_sorce = "a:swCell";
                                } else if (tblStylAttrObj["isFrstRowAttr"] == 1 && i == 0 && getTextByPathList(thisTblStyle, ["a:nwCell"]) !== undefined) {
                                    a_sorce = "a:nwCell";
                                }
                            }

                            let cellParmAry = await this.getTableCellParams(tcNodes[j], getColsGrid, i, j, thisTblStyle, a_sorce)
                            let text = cellParmAry[0];
                            let colStyl = cellParmAry[1];
                            let cssName = cellParmAry[2];
                            let rowSpan = cellParmAry[3];
                            let colSpan = cellParmAry[4];



                            if (rowSpan !== undefined) {
                                totalrowSpan++;
                                rowSpanAry[j] = parseInt(rowSpan) - 1;
                                tableHtml += "<td class='" + cssName + "' data-row='" + i + "," + j + "' rowspan ='" +
                                    parseInt(rowSpan) + "' style='" + colStyl + "'>" + text + "</td>";
                            } else if (colSpan !== undefined) {
                                tableHtml += "<td class='" + cssName + "' data-row='" + i + "," + j + "' colspan = '" +
                                    parseInt(colSpan) + "' style='" + colStyl + "'>" + text + "</td>";
                                totalColSpan = parseInt(colSpan) - 1;
                            } else {
                                tableHtml += "<td class='" + cssName + "' data-row='" + i + "," + j + "' style = '" + colStyl + "'>" + text + "</td>";
                            }

                        } else {
                            if (rowSpanAry[j] != 0) {
                                rowSpanAry[j] -= 1;
                            }
                            if (totalColSpan != 0) {
                                totalColSpan--;
                            }
                        }
                        j++;
                    }
                } else {
                    //single column 

                    let a_sorce;
                    if (tblStylAttrObj["isFrstColAttr"] == 1 && !(tblStylAttrObj["isLstRowAttr"] == 1)) {
                        a_sorce = "a:firstCol";

                    } else if ((tblStylAttrObj["isBandColAttr"] == 1) && !(tblStylAttrObj["isLstRowAttr"] == 1)) {

                        let aBandNode = getTextByPathList(thisTblStyle, ["a:band2V"]);
                        if (aBandNode === undefined) {
                            aBandNode = getTextByPathList(thisTblStyle, ["a:band1V"]);
                            if (aBandNode !== undefined) {
                                a_sorce = "a:band2V";
                            }
                        } else {
                            a_sorce = "a:band2V";
                        }
                    }

                    if (tblStylAttrObj["isLstColAttr"] == 1 && !(tblStylAttrObj["isLstRowAttr"] == 1)) {
                        a_sorce = "a:lastCol";
                    }


                    let cellParmAry = await this.getTableCellParams(tcNodes, getColsGrid, i, undefined, thisTblStyle, a_sorce)
                    let text = cellParmAry[0];
                    let colStyl = cellParmAry[1];
                    let cssName = cellParmAry[2];
                    let rowSpan = cellParmAry[3];

                    if (rowSpan !== undefined) {
                        tableHtml += "<td  class='" + cssName + "' rowspan='" + parseInt(rowSpan) + "' style = '" + colStyl + "'>" + text + "</td>";
                    } else {
                        tableHtml += "<td class='" + cssName + "' style='" + colStyl + "'>" + text + "</td>";
                    }
                }
            }
            tableHtml += "</tr>";
        }
        //////////////////////////////////////////////////////////////////////////////////


        return tableHtml;
    }

    private async getTableCellParams(tcNodes: any, getColsGrid: any, row_idx: number, col_idx: number | undefined, thisTblStyle: any, cellSource: any) {
        //thisTblStyle["a:band1V"] => thisTblStyle[cellSource]
        //text, cell-width, cell-borders, 
        //var text = genTextBody(tcNodes["a:txBody"], tcNodes, undefined, undefined, undefined, undefined, this.resources);//tableStyles
        let rowSpan = getTextByPathList(tcNodes, ["attrs", "rowSpan"]);
        let colSpan = getTextByPathList(tcNodes, ["attrs", "gridSpan"]);
        let vMerge = getTextByPathList(tcNodes, ["attrs", "vMerge"]);
        let hMerge = getTextByPathList(tcNodes, ["attrs", "hMerge"]);
        let colStyl = "word-wrap: break-word;";
        let colWidth;
        let celFillColor = "";
        let col_borders = "";
        let colFontClrPr = "";
        let colFontWeight = "";
        let lin_bottm = "",
            lin_top = "",
            lin_left = "",
            lin_right = "",
            lin_bottom_left_to_top_right = "",
            lin_top_left_to_bottom_right = "";

        let colSapnInt = parseInt(colSpan);
        let total_col_width = 0;
        if (!isNaN(colSapnInt) && colSapnInt > 1) {
            for (var k = 0; k < colSapnInt; k++) {
                total_col_width += parseInt(getTextByPathList(getColsGrid[col_idx! + k], ["attrs", "w"]));
            }
        } else {
            total_col_width = getTextByPathList((col_idx === undefined) ? getColsGrid : getColsGrid[col_idx], ["attrs", "w"]);
        }


        let text = await this.genTextBody(tcNodes["a:txBody"], tcNodes, undefined, undefined, undefined, undefined, total_col_width);//tableStyles

        if (total_col_width != 0 /*&& row_idx == 0*/) {
            colWidth = _.toInteger(total_col_width) * this.options.slideFactor;
            colStyl += "width:" + colWidth + "px;";
        }

        //cell bords
        lin_bottm = getTextByPathList(tcNodes, ["a:tcPr", "a:lnB"]);
        if (lin_bottm === undefined && cellSource !== undefined) {
            if (cellSource !== undefined)
                lin_bottm = getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:bottom", "a:ln"]);
            if (lin_bottm === undefined) {
                lin_bottm = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:bottom", "a:ln"]);
            }
        }
        lin_top = getTextByPathList(tcNodes, ["a:tcPr", "a:lnT"]);
        if (lin_top === undefined) {
            if (cellSource !== undefined)
                lin_top = getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:top", "a:ln"]);
            if (lin_top === undefined) {
                lin_top = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:top", "a:ln"]);
            }
        }
        lin_left = getTextByPathList(tcNodes, ["a:tcPr", "a:lnL"]);
        if (lin_left === undefined) {
            if (cellSource !== undefined)
                lin_left = getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:left", "a:ln"]);
            if (lin_left === undefined) {
                lin_left = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:left", "a:ln"]);
            }
        }
        lin_right = getTextByPathList(tcNodes, ["a:tcPr", "a:lnR"]);
        if (lin_right === undefined) {
            if (cellSource !== undefined)
                lin_right = getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:right", "a:ln"]);
            if (lin_right === undefined) {
                lin_right = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:right", "a:ln"]);
            }
        }
        lin_bottom_left_to_top_right = getTextByPathList(tcNodes, ["a:tcPr", "a:lnBlToTr"]);
        lin_top_left_to_bottom_right = getTextByPathList(tcNodes, ["a:tcPr", "a:InTlToBr"]);

        if (lin_bottm !== undefined && lin_bottm != "") {
            let bottom_line_border = this.getBorder(lin_bottm, undefined, false, "")
            if (bottom_line_border != "") {
                colStyl += "border-bottom:" + bottom_line_border + ";";
            }
        }
        if (lin_top !== undefined && lin_top != "") {
            let top_line_border = this.getBorder(lin_top, undefined, false, "");
            if (top_line_border != "") {
                colStyl += "border-top: " + top_line_border + ";";
            }
        }
        if (lin_left !== undefined && lin_left != "") {
            let left_line_border = this.getBorder(lin_left, undefined, false, "")
            if (left_line_border != "") {
                colStyl += "border-left: " + left_line_border + ";";
            }
        }
        if (lin_right !== undefined && lin_right != "") {
            let right_line_border = this.getBorder(lin_right, undefined, false, "")
            if (right_line_border != "") {
                colStyl += "border-right:" + right_line_border + ";";
            }
        }

        //cell fill color custom
        let getCelFill = getTextByPathList(tcNodes, ["a:tcPr"]);
        if (getCelFill !== undefined && getCelFill != "") {
            let cellObj = {
                "p:spPr": getCelFill
            };
            celFillColor = await this.getShapeFill(cellObj, undefined, false, "slide")
        }

        //cell fill color theme
        if (celFillColor == "" || celFillColor == "background-color: inherit;") {
            let bgFillschemeClr;
            if (cellSource !== undefined)
                bgFillschemeClr = getTextByPathList(thisTblStyle, [cellSource, "a:tcStyle", "a:fill", "a:solidFill"]);
            if (bgFillschemeClr !== undefined) {
                let local_fillColor = this.getSolidFill(bgFillschemeClr);
                if (local_fillColor !== undefined) {
                    celFillColor = " background-color: #" + local_fillColor + ";";
                }
            }
        }
        let cssName = "";
        if (celFillColor !== undefined && celFillColor != "") {
            if (this.options.styleTable?.has(celFillColor)) {
                cssName = this.options.styleTable.get(celFillColor).name;
            } else {
                cssName = "_tbl_cell_css_" + this.options.styleTable?.nextIndex();
                this.options.styleTable?.update(celFillColor, {
                    "name": cssName,
                    "text": celFillColor
                });
            }

        }

        //border
        // let borderStyl = getTextByPathList(thisTblStyle, [cellSource, "a:tcStyle", "a:tcBdr"]);
        // if (borderStyl !== undefined) {
        //     let local_col_borders = this.getTableBorders(borderStyl, this.resources);
        //     if (local_col_borders != "") {
        //         col_borders = local_col_borders;
        //     }
        // }
        // if (col_borders != "") {
        //     colStyl += col_borders;
        // }

        //Text style
        let rowTxtStyl;
        if (cellSource !== undefined) {
            rowTxtStyl = getTextByPathList(thisTblStyle, [cellSource, "a:tcTxStyle"]);
        }
        // if (rowTxtStyl === undefined) {
        //     rowTxtStyl = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
        // }
        if (rowTxtStyl !== undefined) {
            let local_fontClrPr = this.getSolidFill(rowTxtStyl);
            if (local_fontClrPr !== undefined) {
                colFontClrPr = local_fontClrPr;
            }
            let local_fontWeight = ((getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
            if (local_fontWeight !== "") {
                colFontWeight = local_fontWeight;
            }
        }
        colStyl += ((colFontClrPr !== "") ? "color: #" + colFontClrPr + ";" : "");
        colStyl += ((colFontWeight != "") ? " font-weight:" + colFontWeight + ";" : "");

        return [text, colStyl, cssName, rowSpan, colSpan];
    }

    private async genChart(node: any) {

        let order = node["attrs"]?.["order"];
        let xfrmNode = getTextByPathList(node, ["p:xfrm"]);
        const chartID = `chart-${uuidv4()}`;
        let result = "<div id='" + chartID + "' class='block content' style='" +
            this.getPosition(xfrmNode, node, undefined, undefined) + this.getSize(xfrmNode, node, undefined, undefined) +
            " z-index: " + order + ";'></div>";

        let rid = node["a:graphic"]["a:graphicData"]["c:chart"]["attrs"]?.["r:id"];
        let refName = this.resources?.slideResObj[rid]["target"];
        let content = await this.readXmlFile(refName);
        let plotArea = getTextByPathList(content, ["c:chartSpace", "c:chart", "c:plotArea"]);

        let chartData = null;
        for (var key in plotArea) {
            switch (key) {
                case "c:lineChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": chartID,
                            "chartType": "lineChart",
                            "chartData": this.extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:barChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": chartID,
                            "chartType": "barChart",
                            "chartData": this.extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:pieChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": chartID,
                            "chartType": "pieChart",
                            "chartData": this.extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:pie3DChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": chartID,
                            "chartType": "pie3DChart",
                            "chartData": this.extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:areaChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": chartID,
                            "chartType": "areaChart",
                            "chartData": this.extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:scatterChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": chartID,
                            "chartType": "scatterChart",
                            "chartData": this.extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:catAx":
                    break;
                case "c:valAx":
                    break;
                default:
            }
        }

        if (chartData !== null) {
            this.MsgQueue.push(chartData);
        }

        return result;
    }

    private async genDiagram(node: any, source: string, sType: string | undefined) {
        //console.log(this.resources)
        //readXmlFile(zip, sldFileName)
        /**files define the diagram:
         * 1-colors#.xml,
         * 2-data#.xml, 
         * 3-layout#.xml,
         * 4-quickStyle#.xml.
         * 5-drawing#.xml, which Microsoft added as an extension for persisting diagram layout information.
         */
        ///get colors#.xml, data#.xml , layout#.xml , quickStyle#.xml
        let order = node["attrs"]?.["order"];
        let xfrmNode = getTextByPathList(node, ["p:xfrm"]);
        let dgmRelIds = getTextByPathList(node, ["a:graphic", "a:graphicData", "dgm:relIds", "attrs"]);
        //console.log(dgmRelIds)
        let dgmClrFileId = dgmRelIds["r:cs"];
        let dgmDataFileId = dgmRelIds["r:dm"];
        let dgmLayoutFileId = dgmRelIds["r:lo"];
        let dgmQuickStyleFileId = dgmRelIds["r:qs"];
        const slideResObj = this.resources?.slideResObj;
        let dgmClrFileName = slideResObj[dgmClrFileId].target,
            dgmDataFileName = slideResObj[dgmDataFileId].target,
            dgmLayoutFileName = slideResObj[dgmLayoutFileId].target,
            dgmQuickStyleFileName = slideResObj[dgmQuickStyleFileId].target;
        //console.log("dgmClrFileName: " , dgmClrFileName,", dgmDataFileName: ",dgmDataFileName,", dgmLayoutFileName: ",dgmLayoutFileName,", dgmQuickStyleFileName: ",dgmQuickStyleFileName);
        const [
            dgmClr,
            dgmData,
            dgmLayout,
            dgmQuickStyle,
        ] = await Promise.all([
            this.readXmlFile(dgmClrFileName),
            this.readXmlFile(dgmDataFileName),
            this.readXmlFile(dgmLayoutFileName),
            this.readXmlFile(dgmQuickStyleFileName),
        ]);
        //console.log(dgmClr,dgmData,dgmLayout,dgmQuickStyle)
        ///get drawing#.xml
        // let dgmDrwFileName = "";
        // let dataModelExt = getTextByPathList(dgmData, ["dgm:dataModel", "dgm:extLst", "a:ext", "dsp:dataModelExt", "attrs"]);
        // if (dataModelExt !== undefined) {
        //     let dgmDrwFileId = dataModelExt["relId"];
        //     dgmDrwFileName = this.resources["slideResObj"][dgmDrwFileId]["target"];
        // }
        // let dgmDrwFile = "";
        // if (dgmDrwFileName != "") {
        //     dgmDrwFile = readXmlFile(zip, dgmDrwFileName);
        // }
        // let dgmDrwSpArray = getTextByPathList(dgmDrwFile, ["dsp:drawing", "dsp:spTree", "dsp:sp"]);
        //var dgmDrwSpArray = getTextByPathList(this.resources["digramFileContent"], ["dsp:drawing", "dsp:spTree", "dsp:sp"]);
        let dgmDrwSpArray = getTextByPathList(this.resources?.digramFileContent, ["p:drawing", "p:spTree", "p:sp"]);
        let rslt = "";
        if (dgmDrwSpArray !== undefined) {
            let dgmDrwSpArrayLen = dgmDrwSpArray.length;
            for (var i = 0; i < dgmDrwSpArrayLen; i++) {
                let dspSp = dgmDrwSpArray[i];
                // let dspSpObjToStr = JSON.stringify(dspSp);
                // let pSpStr = dspSpObjToStr.replace(/dsp:/g, "p:");
                // let pSpStrToObj = JSON.parse(pSpStr);
                //console.log("pSpStrToObj[" + i + "]: ", pSpStrToObj);
                //rslt += processSpNode(pSpStrToObj, node, this.resources, "diagramBg", sType)
                rslt += this.processSpNode(dspSp, node, "diagramBg", sType)
            }
            // dgmDrwFile: "dsp:"-> "p:"
        }

        return "<div class='block diagram-content' style='" +
            this.getPosition(xfrmNode, node, undefined, undefined, sType) +
            this.getSize(xfrmNode, node, undefined, undefined, sType) +
            "'>" + rslt + "</div>";
    }

    private getPosition(slideSpNode: any, pNode: any, slideLayoutSpNode: any, slideMasterSpNode: any, sType?: string) {
        let off;
        let x = -1, y = -1;

        if (slideSpNode !== undefined) {
            off = slideSpNode["a:off"]["attrs"];
        }

        if (off === undefined && slideLayoutSpNode !== undefined) {
            off = slideLayoutSpNode["a:off"]["attrs"];
        } else if (off === undefined && slideMasterSpNode !== undefined) {
            off = slideMasterSpNode["a:off"]["attrs"];
        }
        if (off === undefined) {
            return "";
        } else {
            x = parseInt(off["x"]) * this.options.slideFactor;
            y = parseInt(off["y"]) * this.options.slideFactor;
            // if (type = "body")
            //     console.log("getPosition: slideSpNode: ", slideSpNode, ", type: ", type, "x: ", x, "offX:", offX, "y:", y, "offY:", offY)
            return (isNaN(x) || isNaN(y)) ? "" : "top:" + y + "px; left:" + x + "px;";
        }

    }

    private getSize(slideSpNode: any, pNode: any, slideLayoutSpNode: any, slideMasterSpNode: any, sType?: string) {
        let ext = undefined;
        let w = -1, h = -1;

        if (slideSpNode !== undefined) {
            ext = slideSpNode["a:ext"]["attrs"];
        } else if (slideLayoutSpNode !== undefined) {
            ext = slideLayoutSpNode["a:ext"]["attrs"];
        } else if (slideMasterSpNode !== undefined) {
            ext = slideMasterSpNode["a:ext"]["attrs"];
        }

        if (ext === undefined) {
            return "";
        } else {
            w = parseInt(ext["cx"]) * this.options.slideFactor;
            h = parseInt(ext["cy"]) * this.options.slideFactor;
            return (isNaN(w) || isNaN(h)) ? "" : "width:" + w + "px; height:" + h + "px;";
        }

    }
    private getVerticalMargins(pNode: any, textBodyNode: any, type: string | undefined, idx: number | undefined) {
        //margin-top ; 
        //a:pPr => a:spcBef => a:spcPts (/100) | a:spcPct (/?)
        //margin-bottom
        //a:pPr => a:spcAft => a:spcPts (/100) | a:spcPct (/?)
        //+
        //a:pPr =>a:lnSpc => a:spcPts (/?) | a:spcPct (/?)
        //console.log("getVerticalMargins ", pNode, type,idx, this.resources)
        //var lstStyle = textBodyNode["a:lstStyle"];
        let lvl: string | number = 1
        let spcBefNode = getTextByPathList(pNode, ["a:pPr", "a:spcBef", "a:spcPts", "attrs", "val"]);
        let spcAftNode = getTextByPathList(pNode, ["a:pPr", "a:spcAft", "a:spcPts", "attrs", "val"]);
        let lnSpcNode = getTextByPathList(pNode, ["a:pPr", "a:lnSpc", "a:spcPct", "attrs", "val"]);
        let lnSpcNodeType = "Pct";
        if (lnSpcNode === undefined) {
            lnSpcNode = getTextByPathList(pNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
            if (lnSpcNode !== undefined) {
                lnSpcNodeType = "Pts";
            }
        }
        let lvlNode = getTextByPathList(pNode, ["a:pPr", "attrs", "lvl"]);
        if (lvlNode !== undefined) {
            lvl = parseInt(lvlNode) + 1;
        }
        let fontSize;
        if (getTextByPathList(pNode, ["a:r"]) !== undefined) {
            let fontSizeStr = this.getFontSize(pNode["a:r"], textBodyNode, undefined, lvl, type);
            if (fontSizeStr != "inherit") {
                fontSize = parseInt(fontSizeStr); //pt
            }
        }
        //var spcBef = "";
        //console.log("getVerticalMargins 1", fontSizeStr, fontSize, lnSpcNode, parseInt(lnSpcNode) / 100000, spcBefNode, spcAftNode)
        // if(spcBefNode !== undefined){
        //     spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "pt;"
        // }
        // else{
        //    //i did not found case with percentage 
        //     spcBefNode = getTextByPathList(pNode, ["a:pPr", "a:spcBef", "a:spcPct","attrs","val"]);
        //     if(spcBefNode !== undefined){
        //         spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "%;"
        //     }
        // }
        //var spcAft = "";
        // if(spcAftNode !== undefined){
        //     spcAft = "margin-bottom:" + parseInt(spcAftNode)/100 + "pt;"
        // }
        // else{
        //    //i did not found case with percentage 
        //     spcAftNode = getTextByPathList(pNode, ["a:pPr", "a:spcAft", "a:spcPct","attrs","val"]);
        //     if(spcAftNode !== undefined){
        //         spcBef = "margin-bottom:" + parseInt(spcAftNode)/100 + "%;"
        //     }
        // }
        // if(spcAftNode !== undefined){
        //     //check in layout and then in master
        // }
        let isInLayoutOrMaster = true;
        if (type == "shape" || type == "textBox") {
            isInLayoutOrMaster = false;
        }
        if (isInLayoutOrMaster && (spcBefNode === undefined || spcAftNode === undefined || lnSpcNode === undefined)) {
            //check in layout
            if (idx !== undefined) {
                let laypPrNode = getTextByPathList(this.resources?.slideLayoutTables, ["idxTable", idx, "p:txBody", "a:p", (lvl - 1), "a:pPr"]);

                if (spcBefNode === undefined) {
                    spcBefNode = getTextByPathList(laypPrNode, ["a:spcBef", "a:spcPts", "attrs", "val"]);
                    // if(spcBefNode !== undefined){
                    //     spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "pt;"
                    // } 
                    // else{
                    //    //i did not found case with percentage 
                    //     spcBefNode = getTextByPathList(laypPrNode, ["a:spcBef", "a:spcPct","attrs","val"]);
                    //     if(spcBefNode !== undefined){
                    //         spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "%;"
                    //     }
                    // }
                }

                if (spcAftNode === undefined) {
                    spcAftNode = getTextByPathList(laypPrNode, ["a:spcAft", "a:spcPts", "attrs", "val"]);
                    // if(spcAftNode !== undefined){
                    //     spcAft = "margin-bottom:" + parseInt(spcAftNode)/100 + "pt;"
                    // }
                    // else{
                    //    //i did not found case with percentage 
                    //     spcAftNode = getTextByPathList(laypPrNode, ["a:spcAft", "a:spcPct","attrs","val"]);
                    //     if(spcAftNode !== undefined){
                    //         spcBef = "margin-bottom:" + parseInt(spcAftNode)/100 + "%;"
                    //     }
                    // }
                }

                if (lnSpcNode === undefined) {
                    lnSpcNode = getTextByPathList(laypPrNode, ["a:lnSpc", "a:spcPct", "attrs", "val"]);
                    if (lnSpcNode === undefined) {
                        lnSpcNode = getTextByPathList(laypPrNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
                        if (lnSpcNode !== undefined) {
                            lnSpcNodeType = "Pts";
                        }
                    }
                }
            }

        }
        if (isInLayoutOrMaster && (spcBefNode === undefined || spcAftNode === undefined || lnSpcNode === undefined)) {
            //check in master
            //slideMasterTextStyles
            let slideMasterTextStyles = this.resources?.slideMasterTextStyles;
            let dirLoc = "";
            lvl = "a:lvl" + lvl + "pPr";
            switch (type) {
                case "title":
                case "ctrTitle":
                    dirLoc = "p:titleStyle";
                    break;
                case "body":
                case "obj":
                case "dt":
                case "ftr":
                case "sldNum":
                case "textBox":
                    // case "shape":
                    dirLoc = "p:bodyStyle";
                    break;
                case "shape":
                //case "textBox":
                default:
                    dirLoc = "p:otherStyle";
            }
            // if (type == "shape" || type == "textBox") {
            //     lvl = "a:lvl1pPr";
            // }
            let inLvlNode = getTextByPathList(slideMasterTextStyles, [dirLoc, lvl]);
            if (inLvlNode !== undefined) {
                if (spcBefNode === undefined) {
                    spcBefNode = getTextByPathList(inLvlNode, ["a:spcBef", "a:spcPts", "attrs", "val"]);
                    // if(spcBefNode !== undefined){
                    //     spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "pt;"
                    // } 
                    // else{
                    //    //i did not found case with percentage 
                    //     spcBefNode = getTextByPathList(inLvlNode, ["a:spcBef", "a:spcPct","attrs","val"]);
                    //     if(spcBefNode !== undefined){
                    //         spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "%;"
                    //     }
                    // }
                }

                if (spcAftNode === undefined) {
                    spcAftNode = getTextByPathList(inLvlNode, ["a:spcAft", "a:spcPts", "attrs", "val"]);
                    // if(spcAftNode !== undefined){
                    //     spcAft = "margin-bottom:" + parseInt(spcAftNode)/100 + "pt;"
                    // }
                    // else{
                    //    //i did not found case with percentage 
                    //     spcAftNode = getTextByPathList(inLvlNode, ["a:spcAft", "a:spcPct","attrs","val"]);
                    //     if(spcAftNode !== undefined){
                    //         spcBef = "margin-bottom:" + parseInt(spcAftNode)/100 + "%;"
                    //     }
                    // }
                }

                if (lnSpcNode === undefined) {
                    lnSpcNode = getTextByPathList(inLvlNode, ["a:lnSpc", "a:spcPct", "attrs", "val"]);
                    if (lnSpcNode === undefined) {
                        lnSpcNode = getTextByPathList(inLvlNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
                        if (lnSpcNode !== undefined) {
                            lnSpcNodeType = "Pts";
                        }
                    }
                }
            }
        }
        let spcBefor = 0, spcAfter = 0, spcLines = 0;
        let marginTopBottomStr = "";
        if (spcBefNode !== undefined) {
            spcBefor = parseInt(spcBefNode) / 100;
        }
        if (spcAftNode !== undefined) {
            spcAfter = parseInt(spcAftNode) / 100;
        }

        if (lnSpcNode !== undefined && fontSize !== undefined) {
            if (lnSpcNodeType == "Pts") {
                marginTopBottomStr += "padding-top: " + ((parseInt(lnSpcNode) / 100) - fontSize) + "px;";//+ "pt;";
            } else {
                let fct = parseInt(lnSpcNode) / 100000;
                spcLines = fontSize * (fct - 1) - fontSize;// fontSize *
                let pTop = (fct > 1) ? spcLines : 0;
                let pBottom = (fct > 1) ? fontSize : 0;
                // marginTopBottomStr += "padding-top: " + spcLines + "pt;";
                // marginTopBottomStr += "padding-bottom: " + pBottom + "pt;";
                marginTopBottomStr += "padding-top: " + pBottom + "px;";// + "pt;";
                marginTopBottomStr += "padding-bottom: " + spcLines + "px;";// + "pt;";
            }
        }

        //if (spcBefNode !== undefined || lnSpcNode !== undefined) {
        marginTopBottomStr += "margin-top: " + (spcBefor - 1) + "px;";// + "pt;"; //margin-top: + spcLines // minus 1 - to fix space
        //}
        if (spcAftNode !== undefined || lnSpcNode !== undefined) {
            //marginTopBottomStr += "margin-bottom: " + ((spcAfter - fontSize < 0) ? 0 : (spcAfter - fontSize)) + "pt;"; //margin-bottom: + spcLines
            //marginTopBottomStr += "margin-bottom: " + spcAfter * (1 / 4) + "px;";// + "pt;";
            marginTopBottomStr += "margin-bottom: " + spcAfter + "px;";// + "pt;";
        }

        //console.log("getVerticalMargins 2 fontSize:", fontSize, "lnSpcNode:", lnSpcNode, "spcLines:", spcLines, "spcBefor:", spcBefor, "spcAfter:", spcAfter)
        //console.log("getVerticalMargins 3 ", marginTopBottomStr, pNode, this.resources)

        //return spcAft + spcBef;
        return marginTopBottomStr;
    }
    private getHorizontalAlign(node: any, textBodyNode: any, idx: number | undefined, type: string | undefined, prg_dir: string) {
        let algn = getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
        if (algn === undefined) {
            //var layoutMasterNode = getLayoutAndMasterNode(node, idx, type, this.resources);
            // let pPrNodeLaout = layoutMasterNode.nodeLaout;
            // let pPrNodeMaster = layoutMasterNode.nodeMaster;
            let lvlIdx = 1;
            let lvlNode = getTextByPathList(node, ["a:pPr", "attrs", "lvl"]);
            if (lvlNode !== undefined) {
                lvlIdx = parseInt(lvlNode) + 1;
            }
            let lvlStr = "a:lvl" + lvlIdx + "pPr";

            let lstStyle = textBodyNode["a:lstStyle"];
            algn = getTextByPathList(lstStyle, [lvlStr, "attrs", "algn"]);
            const slideLayoutTables = this.resources?.slideLayoutTables;
            const slideMasterTextStyles = this.resources?.slideMasterTextStyles;

            if (algn === undefined && idx !== undefined) {
                //slidelayout
                algn = getTextByPathList(slideLayoutTables["idxTable"][idx], ["p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                if (algn === undefined) {
                    algn = getTextByPathList(slideLayoutTables["idxTable"][idx], ["p:txBody", "a:p", "a:pPr", "attrs", "algn"]);
                    if (algn === undefined) {
                        algn = getTextByPathList(slideLayoutTables["idxTable"][idx], ["p:txBody", "a:p", (lvlIdx - 1), "a:pPr", "attrs", "algn"]);
                    }
                }
            }
            if (algn === undefined) {
                if (type !== undefined) {
                    //slidelayout
                    algn = getTextByPathList(slideLayoutTables, ["typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);

                    if (algn === undefined) {
                        //masterlayout
                        if (type == "title" || type == "ctrTitle") {
                            algn = getTextByPathList(slideMasterTextStyles, ["p:titleStyle", lvlStr, "attrs", "algn"]);
                        } else if (type == "body" || type == "obj" || type == "subTitle") {
                            algn = getTextByPathList(slideMasterTextStyles, ["p:bodyStyle", lvlStr, "attrs", "algn"]);
                        } else if (type == "shape" || type == "diagram") {
                            algn = getTextByPathList(slideMasterTextStyles, ["p:otherStyle", lvlStr, "attrs", "algn"]);
                        } else if (type == "textBox") {
                            algn = getTextByPathList(this.options.basicInfo?.defaultTextStyle, [lvlStr, "attrs", "algn"]);
                        } else {
                            algn = getTextByPathList(this.resources?.slideMasterTables, ["typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                        }
                    }
                } else {
                    algn = getTextByPathList(slideMasterTextStyles, ["p:bodyStyle", lvlStr, "attrs", "algn"]);
                }
            }
        }

        if (algn === undefined) {
            if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                return "h-mid";
            } else if (type == "sldNum") {
                return "h-right";
            }
        }
        if (algn !== undefined) {
            switch (algn) {
                case "l":
                    if (prg_dir == "pregraph-rtl") {
                        //return "h-right";
                        return "h-left-rtl";
                    } else {
                        return "h-left";
                    }
                    break;
                case "r":
                    if (prg_dir == "pregraph-rtl") {
                        //return "h-left";
                        return "h-right-rtl";
                    } else {
                        return "h-right";
                    }
                    break;
                case "ctr":
                    return "h-mid";
                    break;
                case "just":
                case "dist":
                default:
                    return "h-" + algn;
            }
        }
        //return algn === "ctr" ? "h-mid" : algn === "r" ? "h-right" : "h-left";
    }
    private getPregraphDir(node: any, textBodyNode: any, idx: number | undefined, type: string | undefined) {
        let rtl = getTextByPathList(node, ["a:pPr", "attrs", "rtl"]);
        //console.log("getPregraphDir node:", node, "textBodyNode", textBodyNode, "rtl:", rtl, "idx", idx, "type", type, "this.resources", this.resources)


        if (rtl === undefined) {
            let layoutMasterNode = this.getLayoutAndMasterNode(node, idx, type);
            let pPrNodeLaout = layoutMasterNode.nodeLaout;
            let pPrNodeMaster = layoutMasterNode.nodeMaster;
            rtl = getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
            if (rtl === undefined && type != "shape") {
                rtl = getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
            }
        }

        if (rtl == "1") {
            return "pregraph-rtl";
        } else if (rtl == "0") {
            return "pregraph-ltr";
        }
        return "pregraph-inherit";

        // let contentDir = getContentDir(type, this.resources);
        // console.log("getPregraphDir node:", node["a:r"], "rtl:", rtl, "idx", idx, "type", type, "contentDir:", contentDir)

        // if (contentDir == "content"){
        //     return "pregraph-ltr";
        // } else if (contentDir == "content-rtl"){ 
        //     return "pregraph-rtl";
        // }
        // return "";
    }
    private getVerticalAlign(node: any, slideLayoutSpNode: any, slideMasterSpNode: any, type: string) {

        //X, <a:bodyPr anchor="ctr">, <a:bodyPr anchor="b">
        let anchor = getTextByPathList(node, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
        //console.log("getVerticalAlign anchor:", anchor, "slideLayoutSpNode: ", slideLayoutSpNode)
        if (anchor === undefined) {
            //console.log("getVerticalAlign type:", type," node:", node, "slideLayoutSpNode:", slideLayoutSpNode, "slideMasterSpNode:", slideMasterSpNode)
            anchor = getTextByPathList(slideLayoutSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
            if (anchor === undefined) {
                anchor = getTextByPathList(slideMasterSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
                if (anchor === undefined) {
                    //"If this attribute is omitted, then a value of t, or top is implied."
                    anchor = "t";//getTextByPathList(slideMasterSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
                }
            }
        }
        //console.log("getVerticalAlign:", node, slideLayoutSpNode, slideMasterSpNode, type, anchor)
        return (anchor === "ctr") ? "v-mid" : ((anchor === "b") ? "v-down" : "v-up");
    }

    private getContentDir(node: any, type: string) {
        return "content";
        let defRtl = getTextByPathList(node, ["p:txBody", "a:lstStyle", "a:defPPr", "attrs", "rtl"]);
        if (defRtl !== undefined) {
            if (defRtl == "1") {
                return "content-rtl";
            } else if (defRtl == "0") {
                return "content";
            }
        }
        //var lvl1Rtl = getTextByPathList(node, ["p:txBody", "a:lstStyle", "lvl1pPr", "attrs", "rtl"]);
        // if (lvl1Rtl !== undefined) {
        //     if (lvl1Rtl == "1") {
        //         return "content-rtl";
        //     } else if (lvl1Rtl == "0") {
        //         return "content";
        //     }
        // }
        let rtlCol = getTextByPathList(node, ["p:txBody", "a:bodyPr", "attrs", "rtlCol"]);
        if (rtlCol !== undefined) {
            if (rtlCol == "1") {
                return "content-rtl";
            } else if (rtlCol == "0") {
                return "content";
            }
        }
        //console.log("getContentDir node:", node, "rtlCol:", rtlCol)

        if (type === undefined) {
            return "content";
        }
        let slideMasterTextStyles = this.resources?.slideMasterTextStyles;
        let dirLoc = "";

        switch (type) {
            case "title":
            case "ctrTitle":
                dirLoc = "p:titleStyle";
                break;
            case "body":
            case "dt":
            case "ftr":
            case "sldNum":
            case "textBox":
                dirLoc = "p:bodyStyle";
                break;
            case "shape":
                dirLoc = "p:otherStyle";
        }
        if (slideMasterTextStyles !== undefined && dirLoc !== "") {
            let dirVal = getTextByPathList(slideMasterTextStyles[dirLoc], ["a:lvl1pPr", "attrs", "rtl"]);
            if (dirVal == "1") {
                return "content-rtl";
            }
        }
        // else {
        //     if (type == "textBox") {
        //         let dirVal = getTextByPathList(this.resources, ["defaultTextStyle", "a:lvl1pPr", "attrs", "rtl"]);
        //         if (dirVal == "1") {
        //             return "content-rtl";
        //         }
        //     }
        // }
        return "content";
        //console.log("getContentDir() type:", type, "slideMasterTextStyles:", slideMasterTextStyles,"dirNode:",dirVal)
    }

    private getFontType(node: any, type: string | undefined, pFontStyle: any) {
        let typeface = getTextByPathList(node, ["a:rPr", "a:latin", "attrs", "typeface"]);

        if (typeface === undefined) {
            let fontIdx = "";
            let fontGrup = "";
            if (pFontStyle !== undefined) {
                fontIdx = getTextByPathList(pFontStyle, ["attrs", "idx"]);
            }
            let fontSchemeNode = getTextByPathList(this.resources?.themeContent, ["a:theme", "a:themeElements", "a:fontScheme"]);
            if (fontIdx == "") {
                if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                    fontIdx = "major";
                } else {
                    fontIdx = "minor";
                }
            }
            fontGrup = "a:" + fontIdx + "Font";
            typeface = getTextByPathList(fontSchemeNode, [fontGrup, "a:latin", "attrs", "typeface"]);
        }

        return (typeface === undefined) ? "inherit" : typeface;
    }

    private getFontColorPr(node: any, pNode: any, lstStyle: any, pFontStyle: any, lvl: string | number, idx: number | undefined, type: string | undefined) {
        //text border using: text-shadow: -1px 0 black, 0 1px black, 1px 0 black, 0 -1px black;
        //{getFontColor(..) return color} -> getFontColorPr(..) return array[color,textBordr/shadow]
        //https://stackoverflow.com/questions/2570972/css-font-border
        //https://www.w3schools.com/cssref/css3_pr_text-shadow.asp
        //themeContent
        //console.log("getFontColorPr>> type:", type, ", node: ", node)
        let rPrNode = getTextByPathList(node, ["a:rPr"]);
        let filTyp, color, textBordr = "", colorType = "", highlightColor: string | undefined = "";
        //console.log("getFontColorPr type:", type, ", node: ", node, "pNode:", pNode, "pFontStyle:", pFontStyle)
        if (rPrNode !== undefined) {
            filTyp = this.getFillType(rPrNode);
            if (filTyp == "SOLID_FILL") {
                let solidFillNode = rPrNode["a:solidFill"];// getTextByPathList(node, ["a:rPr", "a:solidFill"]);
                color = this.getSolidFill(solidFillNode);
                let highlightNode = rPrNode["a:highlight"];
                if (highlightNode !== undefined) {
                    highlightColor = this.getSolidFill(highlightNode);
                }
                colorType = "solid";
            } else if (filTyp == "PATTERN_FILL") {
                let pattFill = rPrNode["a:pattFill"];// getTextByPathList(node, ["a:rPr", "a:pattFill"]);
                color = this.getPatternFill(pattFill);
                colorType = "pattern";
            } else if (filTyp == "PIC_FILL") {
                color = this.getBgPicFill(rPrNode, "slideBg");
                //color = getPicFill("slideBg", rPrNode["a:blipFill"], this.resources);
                colorType = "pic";
            } else if (filTyp == "GRADIENT_FILL") {
                let shpFill = rPrNode["a:gradFill"];
                color = this.getGradientFill(shpFill);
                colorType = "gradient";
            }
        }
        if (color === undefined && getTextByPathList(lstStyle, ["a:lvl" + lvl + "pPr", "a:defRPr"]) !== undefined) {
            //lstStyle
            let lstStyledefRPr = getTextByPathList(lstStyle, ["a:lvl" + lvl + "pPr", "a:defRPr"]);
            filTyp = this.getFillType(lstStyledefRPr);
            if (filTyp == "SOLID_FILL") {
                let solidFillNode = lstStyledefRPr["a:solidFill"];// getTextByPathList(node, ["a:rPr", "a:solidFill"]);
                color = this.getSolidFill(solidFillNode);
                let highlightNode = lstStyledefRPr["a:highlight"];
                if (highlightNode !== undefined) {
                    highlightColor = this.getSolidFill(highlightNode);
                }
                colorType = "solid";
            } else if (filTyp == "PATTERN_FILL") {
                let pattFill = lstStyledefRPr["a:pattFill"];// getTextByPathList(node, ["a:rPr", "a:pattFill"]);
                color = this.getPatternFill(pattFill);
                colorType = "pattern";
            } else if (filTyp == "PIC_FILL") {
                color = this.getBgPicFill(lstStyledefRPr, "slideBg");
                //color = getPicFill("slideBg", rPrNode["a:blipFill"], this.resources);
                colorType = "pic";
            } else if (filTyp == "GRADIENT_FILL") {
                let shpFill = lstStyledefRPr["a:gradFill"];
                color = this.getGradientFill(shpFill);
                colorType = "gradient";
            }

        }
        if (color === undefined) {
            let sPstyle = getTextByPathList(pNode, ["p:style", "a:fontRef"]);
            if (sPstyle !== undefined) {
                color = this.getSolidFill(sPstyle);
                if (color !== undefined) {
                    colorType = "solid";
                }
                let highlightNode = sPstyle["a:highlight"]; //is "a:highlight" node in 'a:fontRef' ?
                if (highlightNode !== undefined) {
                    highlightColor = this.getSolidFill(highlightNode);
                }
            }
            if (color === undefined) {
                if (pFontStyle !== undefined) {
                    color = this.getSolidFill(pFontStyle);
                    if (color !== undefined) {
                        colorType = "solid";
                    }
                }
            }
        }
        //console.log("getFontColorPr node", node, "colorType: ", colorType,"color: ",color)

        if (color === undefined) {

            let layoutMasterNode = this.getLayoutAndMasterNode(pNode, idx, type);
            let pPrNodeLaout = layoutMasterNode.nodeLaout;
            let pPrNodeMaster = layoutMasterNode.nodeMaster;

            if (pPrNodeLaout !== undefined) {
                let defRpRLaout = getTextByPathList(pPrNodeLaout, ["a:defRPr", "a:solidFill"]);
                if (defRpRLaout !== undefined) {
                    color = this.getSolidFill(defRpRLaout);
                    let highlightNode = getTextByPathList(pPrNodeLaout, ["a:defRPr", "a:highlight"]);
                    if (highlightNode !== undefined) {
                        highlightColor = this.getSolidFill(highlightNode);
                    }
                    colorType = "solid";
                }
            }
            if (color === undefined) {

                if (pPrNodeMaster !== undefined) {
                    let defRprMaster = getTextByPathList(pPrNodeMaster, ["a:defRPr", "a:solidFill"]);
                    if (defRprMaster !== undefined) {
                        color = this.getSolidFill(defRprMaster);
                        let highlightNode = getTextByPathList(pPrNodeMaster, ["a:defRPr", "a:highlight"]);
                        if (highlightNode !== undefined) {
                            highlightColor = this.getSolidFill(highlightNode);
                        }
                        colorType = "solid";
                    }
                }
            }
        }
        let txtEffects = [];
        let txtEffObj: any = {}
        //textBordr
        let txtBrdrNode = getTextByPathList(node, ["a:rPr", "a:ln"]);
        if (txtBrdrNode !== undefined && txtBrdrNode["a:noFill"] === undefined) {
            let txBrd = this.getBorder(node, pNode, false, "text");
            let txBrdAry = (txBrd as string).split(" ");
            //var brdSize = (parseInt(txBrdAry[0].substring(0, txBrdAry[0].indexOf("pt")))) + "px";
            let brdSize = (parseInt(txBrdAry[0].substring(0, txBrdAry[0].indexOf("px")))) + "px";
            let brdClr = txBrdAry[2];
            //var brdTyp = txBrdAry[1]; //not in use
            //console.log("getFontColorPr txBrdAry:", txBrdAry)
            if (colorType == "solid") {
                textBordr = "-" + brdSize + " 0 " + brdClr + ", 0 " + brdSize + " " + brdClr + ", " + brdSize + " 0 " + brdClr + ", 0 -" + brdSize + " " + brdClr;
                // if (oShadowStr != "") {
                //     textBordr += "," + oShadowStr;
                // } else {
                //     textBordr += ";";
                // }
                txtEffects.push(textBordr);
            } else {
                //textBordr = brdSize + " " + brdClr;
                txtEffObj.border = brdSize + " " + brdClr;
            }
        }
        // else {
        //     //if no border but exist/not exist shadow
        //     if (colorType == "solid") {
        //         textBordr = oShadowStr;
        //     } else {
        //         //TODO
        //     }
        // }
        let txtGlowNode = getTextByPathList(node, ["a:rPr", "a:effectLst", "a:glow"]);
        let oGlowStr = "";
        if (txtGlowNode !== undefined) {
            let glowClr = this.getSolidFill(txtGlowNode);
            let rad = (txtGlowNode["attrs"]?.["rad"]) ? (txtGlowNode["attrs"]?.["rad"] * this.options.slideFactor) : 0;
            oGlowStr = "0 0 " + rad + "px #" + glowClr +
                ", 0 0 " + rad + "px #" + glowClr +
                ", 0 0 " + rad + "px #" + glowClr +
                ", 0 0 " + rad + "px #" + glowClr +
                ", 0 0 " + rad + "px #" + glowClr +
                ", 0 0 " + rad + "px #" + glowClr +
                ", 0 0 " + rad + "px #" + glowClr;
            if (colorType == "solid") {
                txtEffects.push(oGlowStr);
            } else {
                // txtEffObj.glow = {
                //     radiuse: rad,
                //     color: glowClr
                // } 
                txtEffects.push(
                    "drop-shadow(0 0 " + rad / 3 + "px #" + glowClr + ") " +
                    "drop-shadow(0 0 " + rad * 2 / 3 + "px #" + glowClr + ") " +
                    "drop-shadow(0 0 " + rad + "px #" + glowClr + ")"
                );
            }
        }
        let txtShadow = getTextByPathList(node, ["a:rPr", "a:effectLst", "a:outerShdw"]);
        let oShadowStr = "";
        if (txtShadow !== undefined) {
            //https://developer.mozilla.org/en-US/docs/Web/CSS/filter-function/drop-shadow()
            //https://stackoverflow.com/questions/60468487/css-text-with-linear-gradient-shadow-and-text-outline
            //https://css-tricks.com/creating-playful-effects-with-css-text-shadows/
            //https://designshack.net/articles/css/12-fun-css-text-shadows-you-can-copy-and-paste/

            let shadowClr = this.getSolidFill(txtShadow);
            let outerShdwAttrs = txtShadow["attrs"] ?? {};
            // algn: "bl"
            // dir: "2640000"
            // dist: "38100"
            // rotWithShape: "0/1" - Specifies whether the shadow rotates with the shape if the shape is rotated.
            //blurRad (Blur Radius) - Specifies the blur radius of the shadow.
            //kx (Horizontal Skew) - Specifies the horizontal skew angle.
            //ky (Vertical Skew) - Specifies the vertical skew angle.
            //sx (Horizontal Scaling Factor) - Specifies the horizontal scaling this.options.slideFactor; negative scaling causes a flip.
            //sy (Vertical Scaling Factor) - Specifies the vertical scaling this.options.slideFactor; negative scaling causes a flip.
            let algn = outerShdwAttrs["algn"];
            let dir = (outerShdwAttrs["dir"]) ? (parseInt(outerShdwAttrs["dir"]) / 60000) : 0;
            let dist = parseInt(outerShdwAttrs["dist"]) * this.options.slideFactor;//(px) //* (3 / 4); //(pt)
            let rotWithShape = outerShdwAttrs["rotWithShape"];
            let blurRad = (outerShdwAttrs["blurRad"]) ? (parseInt(outerShdwAttrs["blurRad"]) * this.options.slideFactor + "px") : "";
            let sx = (outerShdwAttrs["sx"]) ? (parseInt(outerShdwAttrs["sx"]) / 100000) : 1;
            let sy = (outerShdwAttrs["sy"]) ? (parseInt(outerShdwAttrs["sy"]) / 100000) : 1;
            let vx = dist * Math.sin(dir * Math.PI / 180);
            let hx = dist * Math.cos(dir * Math.PI / 180);

            //console.log("getFontColorPr outerShdwAttrs:", outerShdwAttrs, ", shadowClr:", shadowClr, ", algn: ", algn, ",dir: ", dir, ", dist: ", dist, ",rotWithShape: ", rotWithShape, ", color: ", color)

            if (!isNaN(vx) && !isNaN(hx)) {
                oShadowStr = hx + "px " + vx + "px " + blurRad + " #" + shadowClr;// + ";";
                if (colorType == "solid") {
                    txtEffects.push(oShadowStr);
                } else {

                    // txtEffObj.oShadow = {
                    //     hx: hx,
                    //     vx: vx,
                    //     radius: blurRad,
                    //     color: shadowClr
                    // }

                    //txtEffObj.oShadow = hx + "px " + vx + "px " + blurRad + " #" + shadowClr;

                    txtEffects.push("drop-shadow(" + hx + "px " + vx + "px " + blurRad + " #" + shadowClr + ")");
                }
            }
            //console.log("getFontColorPr vx:", vx, ", hx: ", hx, ", sx: ", sx, ", sy: ", sy, ",oShadowStr: ", oShadowStr)
        }
        //console.log("getFontColorPr>>> color:", color)
        // if (color === undefined || color === "FFF") {
        //     color = "#000";
        // } else {
        //     color = "" + color;
        // }
        let text_effcts = "", txt_effects;
        if (colorType == "solid") {
            if (txtEffects.length > 0) {
                text_effcts = txtEffects.join(",");
            }
            txt_effects = text_effcts + ";"
        } else {
            if (txtEffects.length > 0) {
                text_effcts = txtEffects.join(" ");
            }
            txtEffObj.effcts = text_effcts;
            txt_effects = txtEffObj
        }
        //console.log("getFontColorPr txt_effects:", txt_effects)

        //return [color, textBordr, colorType];
        return [color, txt_effects, colorType, highlightColor];
    }
    private getFontSize(node: any, textBodyNode: any, pFontStyle: any, lvl?: string | number, type?: string) {
        // if(type == "sldNum")
        //console.log("getFontSize node:", node, "lstStyle", lstStyle, "lvl:", lvl, 'type:', type, "this.resources:", this.resources)
        let lstStyle = (textBodyNode !== undefined) ? textBodyNode["a:lstStyle"] : undefined;
        let lvlpPr = "a:lvl" + lvl + "pPr";
        let fontSize = undefined;
        let sz, kern;
        if (node["a:rPr"] !== undefined) {
            fontSize = parseInt(node["a:rPr"]["attrs"]?.["sz"]) / 100;
        }
        if (isNaN(fontSize!) || fontSize === undefined && node["a:fld"] !== undefined) {
            sz = getTextByPathList(node["a:fld"], ["a:rPr", "attrs", "sz"]);
            fontSize = parseInt(sz) / 100;
        }
        if ((isNaN(fontSize!) || fontSize === undefined) && node["a:t"] === undefined) {
            sz = getTextByPathList(node["a:endParaRPr"], ["attrs", "sz"]);
            fontSize = parseInt(sz) / 100;
        }
        if ((isNaN(fontSize!) || fontSize === undefined) && lstStyle !== undefined) {
            sz = getTextByPathList(lstStyle, [lvlpPr, "a:defRPr", "attrs", "sz"]);
            fontSize = parseInt(sz) / 100;
        }
        //a:spAutoFit
        let isAutoFit = false;
        let isKerning = false;
        if (textBodyNode !== undefined) {
            let spAutoFitNode = getTextByPathList(textBodyNode, ["a:bodyPr", "a:spAutoFit"]);
            // if (spAutoFitNode === undefined) {
            //     spAutoFitNode = getTextByPathList(textBodyNode, ["a:bodyPr", "a:normAutofit"]);
            // }
            if (spAutoFitNode !== undefined) {
                isAutoFit = true;
                isKerning = true;
            }
        }
        if (isNaN(fontSize!) || fontSize === undefined) {
            // if (type == "shape" || type == "textBox") {
            //     type = "body";
            //     lvlpPr = "a:lvl1pPr";
            // }
            sz = getTextByPathList(this.resources?.slideLayoutTables, ["typeTable", type!, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
            fontSize = parseInt(sz) / 100;
            kern = getTextByPathList(this.resources?.slideLayoutTables, ["typeTable", type!, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
            if (isKerning && kern !== undefined && !isNaN(fontSize) && (fontSize - parseInt(kern) / 100) > 0) {
                fontSize = fontSize - parseInt(kern) / 100;
            }
        }

        if (isNaN(fontSize) || fontSize === undefined) {
            // if (type == "shape" || type == "textBox") {
            //     type = "body";
            //     lvlpPr = "a:lvl1pPr";
            // }
            sz = getTextByPathList(this.resources?.slideMasterTables, ["typeTable", type!, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
            kern = getTextByPathList(this.resources?.slideMasterTables, ["typeTable", type!, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
            if (sz === undefined) {
                const slideMasterTextStyles = this.resources?.slideMasterTextStyles;
                if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                    sz = getTextByPathList(slideMasterTextStyles, ["p:titleStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                    kern = getTextByPathList(slideMasterTextStyles, ["p:titleStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                } else if (type == "body" || type == "obj" || type == "dt" || type == "sldNum" || type === "textBox") {
                    sz = getTextByPathList(slideMasterTextStyles, ["p:bodyStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                    kern = getTextByPathList(slideMasterTextStyles, ["p:bodyStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                }
                else if (type == "shape") {
                    //textBox and shape text does not indent
                    sz = getTextByPathList(slideMasterTextStyles, ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                    kern = getTextByPathList(slideMasterTextStyles, ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                    isKerning = false;
                }

                if (sz === undefined) {
                    sz = getTextByPathList(this.options.basicInfo?.defaultTextStyle, [lvlpPr, "a:defRPr", "attrs", "sz"]);
                    kern = (kern === undefined) ? getTextByPathList(this.options.basicInfo?.defaultTextStyle, [lvlpPr, "a:defRPr", "attrs", "kern"]) : undefined;
                    isKerning = false;
                }
                //  else if (type === undefined || type == "shape") {
                //     sz = getTextByPathList(this.resources["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                //     kern = getTextByPathList(this.resources["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                // } 
                // else if (type == "textBox") {
                //     sz = getTextByPathList(this.resources["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                //     kern = getTextByPathList(this.resources["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                // }
            }
            fontSize = parseInt(sz) / 100;
            if (isKerning && kern !== undefined && !isNaN(fontSize) && ((fontSize - parseInt(kern) / 100) > parseInt(kern) / 100)) {
                fontSize = fontSize - parseInt(kern) / 100;
                //fontSize =  parseInt(kern) / 100;
            }
        }

        let baseline = getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
        if (baseline !== undefined && !isNaN(fontSize)) {
            let baselineVl = parseInt(baseline) / 100000;
            //fontSize -= 10; 
            // fontSize = fontSize * baselineVl;
            fontSize -= baselineVl;
        }

        if (!isNaN(fontSize)) {
            let normAutofit = getTextByPathList(textBodyNode, ["a:bodyPr", "a:normAutofit", "attrs", "fontScale"]);
            if (normAutofit !== undefined && normAutofit != 0) {
                //console.log("fontSize", fontSize, "normAutofit: ", normAutofit, normAutofit/100000)
                fontSize = Math.round(fontSize * (normAutofit / 100000))
            }
        }

        return isNaN(fontSize) ? ((type == "br") ? "initial" : "inherit") : (fontSize * this.options.fontSizeFactor + "px");// + "pt");
    }

    private getFontBold(node: any) {
        return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]?.["b"] === "1") ? "bold" : "inherit";
    }

    private getFontItalic(node: any) {
        return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]?.["i"] === "1") ? "italic" : "inherit";
    }

    private getFontDecoration(node: any) {
        ///////////////////////////////Amir///////////////////////////////
        if (node["a:rPr"] !== undefined) {
            let underLine = node["a:rPr"]["attrs"]?.["u"] !== undefined ? node["a:rPr"]["attrs"]?.["u"] : "none";
            let strikethrough = node["a:rPr"]["attrs"]?.["strike"] !== undefined ? node["a:rPr"]["attrs"]?.["strike"] : 'noStrike';
            //console.log("strikethrough: "+strikethrough);

            if (underLine != "none" && strikethrough == "noStrike") {
                return "underline";
            } else if (underLine == "none" && strikethrough != "noStrike") {
                return "line-through";
            } else if (underLine != "none" && strikethrough != "noStrike") {
                return "underline line-through";
            } else {
                return "inherit";
            }
        } else {
            return "inherit";
        }
        /////////////////////////////////////////////////////////////////
        //return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]?.["u"] === "sng") ? "underline" : "inherit";
    }
    ////////////////////////////////////Amir/////////////////////////////////////
    private getTextHorizontalAlign(node: any, pNode: any, type: string | undefined) {
        //console.log("getTextHorizontalAlign: type: ", type, ", node: ", node)
        let getAlgn = getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
        if (getAlgn === undefined) {
            getAlgn = getTextByPathList(pNode, ["a:pPr", "attrs", "algn"]);
        }
        if (getAlgn === undefined) {
            if (type == "title" || type == "ctrTitle" || type == "subTitle") {
                let lvlIdx = 1;
                let lvlNode = getTextByPathList(pNode, ["a:pPr", "attrs", "lvl"]);
                if (lvlNode !== undefined) {
                    lvlIdx = parseInt(lvlNode) + 1;
                }
                let lvlStr = "a:lvl" + lvlIdx + "pPr";
                getAlgn = getTextByPathList(this.resources, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                if (getAlgn === undefined) {
                    getAlgn = getTextByPathList(this.resources, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                    if (getAlgn === undefined) {
                        getAlgn = getTextByPathList(this.resources, ["slideMasterTextStyles", "p:titleStyle", lvlStr, "attrs", "algn"]);
                        if (getAlgn === undefined && type === "subTitle") {
                            getAlgn = getTextByPathList(this.resources, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
                        }
                    }
                }
            } else if (type == "body") {
                getAlgn = getTextByPathList(this.resources, ["slideMasterTextStyles", "p:bodyStyle", "a:lvl1pPr", "attrs", "algn"]);
            } else {
                getAlgn = getTextByPathList(this.resources, ["slideMasterTables", "typeTable", type!, "p:txBody", "a:lstStyle", "a:lvl1pPr", "attrs", "algn"]);
            }

        }

        let align = "inherit";
        if (getAlgn !== undefined) {
            switch (getAlgn) {
                case "l":
                    align = "left";
                    break;
                case "r":
                    align = "right";
                    break;
                case "ctr":
                    align = "center";
                    break;
                case "just":
                    align = "justify";
                    break;
                case "dist":
                    align = "justify";
                    break;
                default:
                    align = "inherit";
            }
        }
        return align;
    }
    /////////////////////////////////////////////////////////////////////
    private getTextVerticalAlign(node: any) {
        let baseline = getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
        return baseline === undefined ? "baseline" : (parseInt(baseline) / 1000) + "%";
    }

    private getTableBorders(node: any) {
        let borderStyle = "";
        if (node["a:bottom"] !== undefined) {
            let obj = {
                "p:spPr": {
                    "a:ln": node["a:bottom"]["a:ln"]
                }
            }
            let borders = this.getBorder(obj, undefined, false, "shape");
            borderStyle += (borders as string).replace("border", "border-bottom");
        }
        if (node["a:top"] !== undefined) {
            let obj = {
                "p:spPr": {
                    "a:ln": node["a:top"]["a:ln"]
                }
            }
            let borders = this.getBorder(obj, undefined, false, "shape");
            borderStyle += (borders as string).replace("border", "border-top");
        }
        if (node["a:right"] !== undefined) {
            let obj = {
                "p:spPr": {
                    "a:ln": node["a:right"]["a:ln"]
                }
            }
            let borders = this.getBorder(obj, undefined, false, "shape");
            borderStyle += (borders as string).replace("border", "border-right");
        }
        if (node["a:left"] !== undefined) {
            let obj = {
                "p:spPr": {
                    "a:ln": node["a:left"]["a:ln"]
                }
            }
            let borders = this.getBorder(obj, undefined, false, "shape");
            borderStyle += (borders as string).replace("border", "border-left");
        }

        return borderStyle;
    }
    //////////////////////////////////////////////////////////////////
    private getBorder(node: any, pNode: any, isSvgMode: boolean, bType: string | undefined) {
        //console.log("getBorder", node, pNode, isSvgMode, bType)
        let cssText, lineNode, subNodeTxt, borderWidth, borderType, strokeDasharray = "0";

        if (bType == "shape") {
            cssText = "border: ";
            lineNode = node["p:spPr"]["a:ln"];
            //subNodeTxt = "p:spPr";
            //node["p:style"]["a:lnRef"] = 
        } else if (bType == "text") {
            cssText = "";
            lineNode = node["a:rPr"]["a:ln"];
            //subNodeTxt = "a:rPr";
        }

        //console.log("lineNode: ", lineNode)
        if (lineNode == undefined) {
            let lnRefNode = getTextByPathList(node, ["p:style", "a:lnRef"])
            if (lnRefNode !== undefined) {
                let lnIdx = getTextByPathList(lnRefNode, ["attrs", "idx"]);
                //console.log("lnIdx:", lnIdx, "lnStyleLst:", this.resources["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:lnStyleLst"]["a:ln"][Number(lnIdx) -1])
                lineNode = this.resources?.["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:lnStyleLst"]["a:ln"][Number(lnIdx) - 1];
            }
        }
        if (lineNode == undefined) {
            //is table
            cssText = "";
            lineNode = node
        }

        //var is_noFill = getTextByPathList(node, ["p:spPr", "a:noFill"]);
        let is_noFill = getTextByPathList(lineNode, ["a:noFill"]);
        if (is_noFill !== undefined) {
            return "hidden";
        }

        let borderColor;
        if (lineNode !== undefined) {
            // Border width: 1pt = 12700, default = 0.75pt
            borderWidth = parseInt(getTextByPathList(lineNode, ["attrs", "w"])) / 12700;
            if (isNaN(borderWidth) || borderWidth < 1) {
                cssText += (4 / 3) + "px ";//"1pt ";
            } else {
                cssText += borderWidth + "px ";// + "pt ";
            }
            // Border type
            borderType = getTextByPathList(lineNode, ["a:prstDash", "attrs", "val"]);
            if (borderType === undefined) {
                borderType = getTextByPathList(lineNode, ["attrs", "cmpd"]);
            }
            switch (borderType) {
                case "solid":
                    cssText += "solid";
                    strokeDasharray = "0";
                    break;
                case "dash":
                    cssText += "dashed";
                    strokeDasharray = "5";
                    break;
                case "dashDot":
                    cssText += "dashed";
                    strokeDasharray = "5, 5, 1, 5";
                    break;
                case "dot":
                    cssText += "dotted";
                    strokeDasharray = "1, 5";
                    break;
                case "lgDash":
                    cssText += "dashed";
                    strokeDasharray = "10, 5";
                    break;
                case "dbl":
                    cssText += "double";
                    strokeDasharray = "0";
                    break;
                case "lgDashDotDot":
                    cssText += "dashed";
                    strokeDasharray = "10, 5, 1, 5, 1, 5";
                    break;
                case "sysDash":
                    cssText += "dashed";
                    strokeDasharray = "5, 2";
                    break;
                case "sysDashDot":
                    cssText += "dashed";
                    strokeDasharray = "5, 2, 1, 5";
                    break;
                case "sysDashDotDot":
                    cssText += "dashed";
                    strokeDasharray = "5, 2, 1, 5, 1, 5";
                    break;
                case "sysDot":
                    cssText += "dotted";
                    strokeDasharray = "2, 5";
                    break;
                case undefined:
                //console.log(borderType);
                default:
                    cssText += "solid";
                    strokeDasharray = "0";
            }
            // Border color
            let fillTyp = this.getFillType(lineNode);
            //console.log("getBorder:node : fillTyp", fillTyp)
            if (fillTyp == "NO_FILL") {
                borderColor = isSvgMode ? "none" : "";//"background-color: initial;";
            } else if (fillTyp == "SOLID_FILL") {
                borderColor = this.getSolidFill(lineNode["a:solidFill"]);
            } else if (fillTyp == "GRADIENT_FILL") {
                borderColor = this.getGradientFill(lineNode["a:gradFill"]);
                //console.log("shpFill",shpFill,grndColor.color)
            } else if (fillTyp == "PATTERN_FILL") {
                borderColor = this.getPatternFill(lineNode["a:pattFill"]);
            }

        }

        //console.log("getBorder:node : borderColor", borderColor)
        // 2. drawingML namespace
        if (borderColor === undefined) {
            //var schemeClrNode = getTextByPathList(node, ["p:style", "a:lnRef", "a:schemeClr"]);
            // if (schemeClrNode !== undefined) {
            //     let schemeClr = "a:" + getTextByPathList(schemeClrNode, ["attrs", "val"]);
            //     let borderColor = getSchemeColorFromTheme(schemeClr, undefined, undefined);
            // }
            let lnRefNode = getTextByPathList(node, ["p:style", "a:lnRef"]);
            //console.log("getBorder: lnRef : ", lnRefNode)
            if (lnRefNode !== undefined) {
                borderColor = this.getSolidFill(lnRefNode);
            }

            // if (borderColor !== undefined) {
            //     let shade = getTextByPathList(schemeClrNode, ["a:shade", "attrs", "val"]);
            //     if (shade !== undefined) {
            //         shade = parseInt(shade) / 10000;
            //         let color = tinycolor("#" + borderColor);
            //         borderColor = color.darken(shade).toHex8();//.replace("#", "");
            //     }
            // }

        }

        //console.log("getBorder: borderColor : ", borderColor)
        if (borderColor === undefined) {
            if (isSvgMode) {
                borderColor = "none";
            } else {
                borderColor = "hidden";
            }
        } else {
            borderColor = "#" + borderColor; //wrong if not solid fill - TODO

        }
        cssText += " " + borderColor + " ";//wrong if not solid fill - TODO

        if (isSvgMode) {
            return { "color": borderColor, "width": borderWidth, "type": borderType, "strokeDasharray": strokeDasharray };
        } else {
            return cssText + ";";
        }
        // } else {
        //     if (isSvgMode) {
        //         return { "color": 'none', "width": '0', "type": 'none', "strokeDasharray": '0' };
        //     } else {
        //         return "hidden";
        //     }
        // }
    }
    private async getBackground() {
        //var rslt = "";
        let slideContent = this.resources?.slideContent;
        let slideLayoutContent = this.resources?.slideLayoutContent;
        let slideMasterContent = this.resources?.slideMasterContent;

        let nodesSldLayout = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:spTree"]);
        let nodesSldMaster = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:spTree"]);
        // console.log("slideContent : ", slideContent)
        // console.log("slideLayoutContent : ", slideLayoutContent)
        // console.log("slideMasterContent : ", slideMasterContent)
        //console.log("this.resources : ", this.resources)
        let showMasterSp = getTextByPathList(slideLayoutContent, ["p:sldLayout", "attrs", "showMasterSp"]);
        //console.log("slideLayoutContent : ", slideLayoutContent, ", showMasterSp: ", showMasterSp)
        let bgColor = await this.getSlideBackgroundFill();
        let result = "<div class='slide-background-" + this.options.index + "' style='width:" + this.options.basicInfo!.width + "px; height:" + this.options.basicInfo!.height + "px;" + bgColor + "'>"
        let node_ph_type_ary = [];
        if (nodesSldLayout !== undefined) {
            for (var nodeKey in nodesSldLayout) {
                if (nodesSldLayout[nodeKey].constructor === Array) {
                    for (var i = 0; i < nodesSldLayout[nodeKey].length; i++) {
                        let ph_type = getTextByPathList(nodesSldLayout[nodeKey][i], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                        // if (ph_type !== undefined && ph_type != "pic") {
                        //     node_ph_type_ary.push(ph_type);
                        // }
                        if (ph_type != "pic") {
                            result += await this.processNodesInSlide(nodeKey, nodesSldLayout[nodeKey][i], nodesSldLayout, "slideLayoutBg"); //slideLayoutBg , slideMasterBg
                        }
                    }
                } else {
                    let ph_type = getTextByPathList(nodesSldLayout[nodeKey], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                    // if (ph_type !== undefined && ph_type != "pic") {
                    //     node_ph_type_ary.push(ph_type);
                    // }
                    if (ph_type != "pic") {
                        result += await this.processNodesInSlide(nodeKey, nodesSldLayout[nodeKey], nodesSldLayout, "slideLayoutBg"); //slideLayoutBg, slideMasterBg
                    }
                }
            }
        }
        if (nodesSldMaster !== undefined && (showMasterSp == "1" || showMasterSp === undefined)) {
            for (var nodeKey in nodesSldMaster) {
                if (nodesSldMaster[nodeKey].constructor === Array) {
                    for (var i = 0; i < nodesSldMaster[nodeKey].length; i++) {
                        let ph_type = getTextByPathList(nodesSldMaster[nodeKey][i], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                        //if (node_ph_type_ary.indexOf(ph_type) > -1) {
                        result += await this.processNodesInSlide(nodeKey, nodesSldMaster[nodeKey][i], nodesSldMaster, "slideMasterBg"); //slideLayoutBg , slideMasterBg
                        //}
                    }
                } else {
                    let ph_type = getTextByPathList(nodesSldMaster[nodeKey], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                    //if (node_ph_type_ary.indexOf(ph_type) > -1) {
                    result += await this.processNodesInSlide(nodeKey, nodesSldMaster[nodeKey], nodesSldMaster, "slideMasterBg"); //slideLayoutBg, slideMasterBg
                    //}
                }
            }
        }
        result += "</div>";
        return result;

    }
    private getSlideBackgroundFill() {
        let slideContent = this.resources?.slideContent;
        let slideLayoutContent = this.resources?.slideLayoutContent;
        let slideMasterContent = this.resources?.slideMasterContent;

        //console.log("slideContent: ", slideContent)
        //console.log("slideLayoutContent: ", slideLayoutContent)
        //console.log("slideMasterContent: ", slideMasterContent)
        //getFillType(node)
        let bgPr = getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg", "p:bgPr"]);
        let bgRef = getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg", "p:bgRef"]);
        //console.log("slideContent >> bgPr: ", bgPr, ", bgRef: ", bgRef)
        let bgcolor;
        if (bgPr !== undefined) {
            //bgcolor = "background-color: blue;";
            let bgFillTyp = this.getFillType(bgPr);

            if (bgFillTyp == "SOLID_FILL") {
                let sldFill = bgPr["a:solidFill"];
                let clrMapOvr;
                let sldClrMapOvr = getTextByPathList(slideContent, ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                if (sldClrMapOvr !== undefined) {
                    clrMapOvr = sldClrMapOvr;
                } else {
                    let sldClrMapOvr = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                    if (sldClrMapOvr !== undefined) {
                        clrMapOvr = sldClrMapOvr;
                    } else {
                        clrMapOvr = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
                    }

                }
                let sldBgClr = this.getSolidFill(sldFill, clrMapOvr);
                //var sldTint = getColorOpacity(sldFill);
                //console.log("bgColor: ", bgColor)
                //bgcolor = "background: rgba(" + hexToRgbNew(bgColor) + "," + sldTint + ");";
                bgcolor = "background: #" + sldBgClr + ";";

            } else if (bgFillTyp == "GRADIENT_FILL") {
                bgcolor = this.getBgGradientFill(bgPr);
            } else if (bgFillTyp == "PIC_FILL") {
                //console.log("PIC_FILL - ", bgFillTyp, bgPr, this.resources);
                bgcolor = this.getBgPicFill(bgPr, "slideBg");

            }
            //console.log(slideContent,slideMasterContent,color_ary,tint_ary,rot,bgcolor)
        } else if (bgRef !== undefined) {
            //console.log("slideContent",bgRef)
            let clrMapOvr;
            let sldClrMapOvr = getTextByPathList(slideContent, ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
            if (sldClrMapOvr !== undefined) {
                clrMapOvr = sldClrMapOvr;
            } else {
                let sldClrMapOvr = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                if (sldClrMapOvr !== undefined) {
                    clrMapOvr = sldClrMapOvr;
                } else {
                    clrMapOvr = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
                }

            }
            let phClr = this.getSolidFill(bgRef, clrMapOvr);

            // if (bgRef["a:srgbClr"] !== undefined) {
            //     phClr = getTextByPathList(bgRef, ["a:srgbClr", "attrs", "val"]); //#...
            // } else if (bgRef["a:schemeClr"] !== undefined) { //a:schemeClr
            //     let schemeClr = getTextByPathList(bgRef, ["a:schemeClr", "attrs", "val"]);
            //     phClr = getSchemeColorFromTheme("a:" + schemeClr, slideMasterContent, undefined); //#...
            // }
            let idx = Number(bgRef["attrs"]?.["idx"]);


            if (idx == 0 || idx == 1000) {
                //no background
            } else if (idx > 0 && idx < 1000) {
                //fillStyleLst in themeContent
                //themeContent["a:fmtScheme"]["a:fillStyleLst"]
                //bgcolor = "background: red;";
            } else if (idx > 1000) {
                //bgFillStyleLst  in themeContent
                //themeContent["a:fmtScheme"]["a:bgFillStyleLst"]
                let trueIdx = idx - 1000;
                // themeContent["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                let bgFillLst = this.resources?.["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                let sortblAry: any[] = [];
                Object.keys(bgFillLst).forEach((key) => {
                    let bgFillLstTyp = bgFillLst[key];
                    if (key != "attrs") {
                        if (bgFillLstTyp.constructor === Array) {
                            for (var i = 0; i < bgFillLstTyp.length; i++) {
                                let obj: any = {};
                                obj[key] = bgFillLstTyp[i];
                                obj["idex"] = bgFillLstTyp[i]["attrs"]?.["order"];
                                obj["attrs"] = {
                                    "order": bgFillLstTyp[i]["attrs"]?.["order"]
                                }
                                sortblAry.push(obj)
                            }
                        } else {
                            let obj: any = {};
                            obj[key] = bgFillLstTyp;
                            obj["idex"] = bgFillLstTyp["attrs"]?.["order"];
                            obj["attrs"] = {
                                "order": bgFillLstTyp["attrs"]?.["order"]
                            }
                            sortblAry.push(obj)
                        }
                    }
                });
                let sortByOrder = sortblAry.slice(0);
                sortByOrder.sort(function (a, b) {
                    return a.idex - b.idex;
                });
                let bgFillLstIdx = sortByOrder[trueIdx - 1];
                let bgFillTyp = this.getFillType(bgFillLstIdx);
                if (bgFillTyp == "SOLID_FILL") {
                    let sldFill = bgFillLstIdx["a:solidFill"];
                    let sldBgClr = this.getSolidFill(sldFill, clrMapOvr);
                    //var sldTint = getColorOpacity(sldFill);
                    //bgcolor = "background: rgba(" + hexToRgbNew(phClr) + "," + sldTint + ");";
                    bgcolor = "background: #" + sldBgClr + ";";
                    //console.log("slideMasterContent - sldFill",sldFill)
                } else if (bgFillTyp == "GRADIENT_FILL") {
                    bgcolor = this.getBgGradientFill(bgFillLstIdx, phClr);
                } else {
                    console.log(bgFillTyp)
                }
            }

        }
        else {
            bgPr = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgPr"]);
            bgRef = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgRef"]);
            //console.log("slideLayoutContent >> bgPr: ", bgPr, ", bgRef: ", bgRef)
            let clrMapOvr;
            let sldClrMapOvr = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
            if (sldClrMapOvr !== undefined) {
                clrMapOvr = sldClrMapOvr;
            } else {
                clrMapOvr = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
            }
            if (bgPr !== undefined) {
                let bgFillTyp = this.getFillType(bgPr);
                if (bgFillTyp == "SOLID_FILL") {
                    let sldFill = bgPr["a:solidFill"];

                    let sldBgClr = this.getSolidFill(sldFill, clrMapOvr);
                    //var sldTint = getColorOpacity(sldFill);
                    // bgcolor = "background: rgba(" + hexToRgbNew(bgColor) + "," + sldTint + ");";
                    bgcolor = "background: #" + sldBgClr + ";";
                } else if (bgFillTyp == "GRADIENT_FILL") {
                    bgcolor = this.getBgGradientFill(bgPr);
                } else if (bgFillTyp == "PIC_FILL") {
                    bgcolor = this.getBgPicFill(bgPr, "slideLayoutBg");

                }
                //console.log("slideLayoutContent",bgcolor)
            } else if (bgRef !== undefined) {
                console.log("slideLayoutContent: bgRef", bgRef)
                //bgcolor = "background: white;";
                let phClr = this.getSolidFill(bgRef, clrMapOvr);
                let idx = Number(bgRef["attrs"]?.["idx"]);
                //console.log("phClr=", phClr, "idx=", idx)

                if (idx == 0 || idx == 1000) {
                    //no background
                } else if (idx > 0 && idx < 1000) {
                    //fillStyleLst in themeContent
                    //themeContent["a:fmtScheme"]["a:fillStyleLst"]
                    //bgcolor = "background: red;";
                } else if (idx > 1000) {
                    //bgFillStyleLst  in themeContent
                    //themeContent["a:fmtScheme"]["a:bgFillStyleLst"]
                    let trueIdx = idx - 1000;
                    let bgFillLst = this.resources?.["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                    let sortblAry: any[] = [];
                    Object.keys(bgFillLst).forEach((key) => {
                        //console.log("cubicBezTo[" + key + "]:");
                        let bgFillLstTyp = bgFillLst[key];
                        if (key != "attrs") {
                            if (bgFillLstTyp.constructor === Array) {
                                for (var i = 0; i < bgFillLstTyp.length; i++) {
                                    let obj: any = {};
                                    obj[key] = bgFillLstTyp[i];
                                    obj["idex"] = bgFillLstTyp[i]["attrs"]?.["order"];
                                    obj["attrs"] = {
                                        "order": bgFillLstTyp[i]["attrs"]?.["order"]
                                    }
                                    sortblAry.push(obj)
                                }
                            } else {
                                let obj: any = {};
                                obj[key] = bgFillLstTyp;
                                obj["idex"] = bgFillLstTyp["attrs"]?.["order"];
                                obj["attrs"] = {
                                    "order": bgFillLstTyp["attrs"]?.["order"]
                                }
                                sortblAry.push(obj)
                            }
                        }
                    });
                    let sortByOrder = sortblAry.slice(0);
                    sortByOrder.sort(function (a, b) {
                        return a.idex - b.idex;
                    });
                    let bgFillLstIdx = sortByOrder[trueIdx - 1];
                    let bgFillTyp = this.getFillType(bgFillLstIdx);
                    if (bgFillTyp == "SOLID_FILL") {
                        let sldFill = bgFillLstIdx["a:solidFill"];
                        //console.log("sldFill: ", sldFill)
                        //var sldTint = getColorOpacity(sldFill);
                        //bgcolor = "background: rgba(" + hexToRgbNew(phClr) + "," + sldTint + ");";
                        let sldBgClr = this.getSolidFill(sldFill, clrMapOvr, phClr);
                        //console.log("bgcolor: ", bgcolor)
                        bgcolor = "background: #" + sldBgClr + ";";
                    } else if (bgFillTyp == "GRADIENT_FILL") {
                        //console.log("GRADIENT_FILL: ", bgFillLstIdx, phClr)
                        bgcolor = this.getBgGradientFill(bgFillLstIdx, phClr);
                    } else if (bgFillTyp == "PIC_FILL") {
                        //theme rels
                        //console.log("PIC_FILL - ", bgFillTyp, bgFillLstIdx, bgFillLst, this.resources);
                        bgcolor = this.getBgPicFill(bgFillLstIdx, "themeBg", phClr);
                    } else {
                        console.log(bgFillTyp)
                    }
                }
            } else {
                bgPr = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:bg", "p:bgPr"]);
                bgRef = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:bg", "p:bgRef"]);

                let clrMap = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
                //console.log("slideMasterContent >> bgPr: ", bgPr, ", bgRef: ", bgRef)
                if (bgPr !== undefined) {
                    let bgFillTyp = this.getFillType(bgPr);
                    if (bgFillTyp == "SOLID_FILL") {
                        let sldFill = bgPr["a:solidFill"];
                        let sldBgClr = this.getSolidFill(sldFill, clrMap);
                        // let sldTint = getColorOpacity(sldFill);
                        // bgcolor = "background: rgba(" + hexToRgbNew(bgColor) + "," + sldTint + ");";
                        bgcolor = "background: #" + sldBgClr + ";";
                    } else if (bgFillTyp == "GRADIENT_FILL") {
                        bgcolor = this.getBgGradientFill(bgPr);
                    } else if (bgFillTyp == "PIC_FILL") {
                        bgcolor = this.getBgPicFill(bgPr, "slideMasterBg");
                    }
                } else if (bgRef !== undefined) {
                    //var obj={
                    //    "a:solidFill": bgRef
                    //}
                    let phClr = this.getSolidFill(bgRef, clrMap);
                    // let phClr;
                    // if (bgRef["a:srgbClr"] !== undefined) {
                    //     phClr = getTextByPathList(bgRef, ["a:srgbClr", "attrs", "val"]); //#...
                    // } else if (bgRef["a:schemeClr"] !== undefined) { //a:schemeClr
                    //     let schemeClr = getTextByPathList(bgRef, ["a:schemeClr", "attrs", "val"]);

                    //     phClr = getSchemeColorFromTheme("a:" + schemeClr, slideMasterContent, undefined); //#...
                    // }
                    let idx = Number(bgRef["attrs"]?.["idx"]);
                    //console.log("phClr=", phClr, "idx=", idx)

                    if (idx == 0 || idx == 1000) {
                        //no background
                    } else if (idx > 0 && idx < 1000) {
                        //fillStyleLst in themeContent
                        //themeContent["a:fmtScheme"]["a:fillStyleLst"]
                        //bgcolor = "background: red;";
                    } else if (idx > 1000) {
                        //bgFillStyleLst  in themeContent
                        //themeContent["a:fmtScheme"]["a:bgFillStyleLst"]
                        let trueIdx = idx - 1000;
                        let bgFillLst = this.resources?.["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                        let sortblAry: any[] = [];
                        Object.keys(bgFillLst).forEach((key) => {
                            //console.log("cubicBezTo[" + key + "]:");
                            let bgFillLstTyp = bgFillLst[key];
                            if (key != "attrs") {
                                if (bgFillLstTyp.constructor === Array) {
                                    for (var i = 0; i < bgFillLstTyp.length; i++) {
                                        let obj: any = {};
                                        obj[key] = bgFillLstTyp[i];
                                        obj["idex"] = bgFillLstTyp[i]["attrs"]?.["order"];
                                        obj["attrs"] = {
                                            "order": bgFillLstTyp[i]["attrs"]?.["order"]
                                        }
                                        sortblAry.push(obj)
                                    }
                                } else {
                                    let obj: any = {};
                                    obj[key] = bgFillLstTyp;
                                    obj["idex"] = bgFillLstTyp["attrs"]?.["order"];
                                    obj["attrs"] = {
                                        "order": bgFillLstTyp["attrs"]?.["order"]
                                    }
                                    sortblAry.push(obj)
                                }
                            }
                        });
                        let sortByOrder = sortblAry.slice(0);
                        sortByOrder.sort(function (a, b) {
                            return a.idex - b.idex;
                        });
                        let bgFillLstIdx = sortByOrder[trueIdx - 1];
                        let bgFillTyp = this.getFillType(bgFillLstIdx);
                        //console.log("bgFillLstIdx: ", bgFillLstIdx, ", bgFillTyp: ", bgFillTyp, ", phClr: ", phClr);
                        if (bgFillTyp == "SOLID_FILL") {
                            let sldFill = bgFillLstIdx["a:solidFill"];
                            //console.log("sldFill: ", sldFill)
                            //var sldTint = getColorOpacity(sldFill);
                            //bgcolor = "background: rgba(" + hexToRgbNew(phClr) + "," + sldTint + ");";
                            let sldBgClr = this.getSolidFill(sldFill, clrMap, phClr);
                            //console.log("bgcolor: ", bgcolor)
                            bgcolor = "background: #" + sldBgClr + ";";
                        } else if (bgFillTyp == "GRADIENT_FILL") {
                            //console.log("GRADIENT_FILL: ", bgFillLstIdx, phClr)
                            bgcolor = this.getBgGradientFill(bgFillLstIdx, phClr);
                        } else if (bgFillTyp == "PIC_FILL") {
                            //theme rels
                            // console.log("PIC_FILL - ", bgFillTyp, bgFillLstIdx, bgFillLst, this.resources);
                            bgcolor = this.getBgPicFill(bgFillLstIdx, "themeBg", phClr);
                        } else {
                            console.log(bgFillTyp)
                        }
                    }
                }
            }
        }

        //console.log("bgcolor: ", bgcolor)
        return bgcolor;
    }
    private getBgGradientFill(bgPr: any, phClr?: string) {
        let bgcolor = "";
        const slideMasterContent = this.resources?.slideMasterContent;
        if (bgPr !== undefined) {
            let grdFill = bgPr["a:gradFill"];
            let gsLst = grdFill["a:gsLst"]["a:gs"];
            //var startColorNode, endColorNode;
            let color_ary = [];
            let pos_ary = [];
            //var tint_ary = [];
            for (var i = 0; i < gsLst.length; i++) {
                let lo_tint;
                let lo_color = this.getSolidFill(gsLst[i], slideMasterContent["p:sldMaster"]["p:clrMap"]["attrs"], phClr);
                let pos = getTextByPathList(gsLst[i], ["attrs", "pos"])
                //console.log("pos: ", pos)
                if (pos !== undefined) {
                    pos_ary[i] = pos / 1000 + "%";
                } else {
                    pos_ary[i] = "";
                }
                //console.log("lo_color", lo_color)
                color_ary[i] = "#" + lo_color;
                //tint_ary[i] = (lo_tint !== undefined) ? parseInt(lo_tint) / 100000 : 1;
            }
            //get rot
            let lin = grdFill["a:lin"];
            let rot = 90;
            if (lin !== undefined) {
                rot = angleToDegrees(lin["attrs"]?.["ang"]);// + 270;
                //console.log("rot: ", rot)
                rot = rot + 90;
            }
            bgcolor = "background: linear-gradient(" + rot + "deg,";
            for (var i = 0; i < gsLst.length; i++) {
                if (i == gsLst.length - 1) {
                    //if (phClr === undefined) {
                    //bgcolor += "rgba(" + hexToRgbNew(color_ary[i]) + "," + tint_ary[i] + ")" + ");";
                    bgcolor += color_ary[i] + " " + pos_ary[i] + ");";
                    //} else {
                    //bgcolor += "rgba(" + hexToRgbNew(phClr) + "," + tint_ary[i] + ")" + ");";
                    // bgcolor += "" + phClr + ";";;
                    //}
                } else {
                    //if (phClr === undefined) {
                    //bgcolor += "rgba(" + hexToRgbNew(color_ary[i]) + "," + tint_ary[i] + ")" + ", ";
                    bgcolor += color_ary[i] + " " + pos_ary[i] + ", ";;
                    //} else {
                    //bgcolor += "rgba(" + hexToRgbNew(phClr) + "," + tint_ary[i] + ")" + ", ";
                    // bgcolor += phClr + ", ";
                    //}
                }
            }
        } else {
            if (phClr !== undefined) {
                //bgcolor = "rgba(" + hexToRgbNew(phClr) + ",0);";
                //bgcolor = phClr + ");";
                bgcolor = "background: #" + phClr + ";";
            }
        }
        return bgcolor;
    }
    private async getBgPicFill(bgPr: any, sorce?: any, phClr?: string) {
        //console.log("getBgPicFill bgPr", bgPr)
        let bgcolor;
        let picFillBase64 = await this.getPicFill(sorce, bgPr["a:blipFill"]);
        let ordr = bgPr["attrs"]?.["order"];
        let aBlipNode = bgPr["a:blipFill"]["a:blip"];
        //a:duotone
        let duotone = getTextByPathList(aBlipNode, ["a:duotone"]);
        if (duotone !== undefined) {
            //console.log("pic duotone: ", duotone)
            let clr_ary = [];
            // duotone.forEach(function (clr) {
            //     console.log("pic duotone clr: ", clr)
            // }) 
            Object.keys(duotone).forEach((clr_type) => {
                //console.log("pic duotone clr: clr_type: ", clr_type, duotone[clr_type])
                if (clr_type != "attrs") {
                    let obj: any = {};
                    obj[clr_type] = duotone[clr_type];
                    clr_ary.push(this.getSolidFill(obj, undefined, phClr));
                }
                // Object.keys(duotone[clr_type]).forEach(function (clr) {
                //     if (clr != "order") {
                //         let obj = {};
                //         obj[clr_type] = duotone[clr_type][clr];
                //         clr_ary.push(getSolidFill(obj, undefined, phClr, this.resources));
                //     }
                // })
            })
            //console.log("pic duotone clr_ary: ", clr_ary);
            //filter: url(file.svg#filter-element-id)
            //https://codepen.io/bhenbe/pen/QEZOvd
            //https://www.w3schools.com/cssref/css3_pr_filter.asp

            // let color1 = clr_ary[0];
            // let color2 = clr_ary[1];
            // let cssName = "";

            // let styleText_before_after = "content: '';" +
            //     "display: block;" +
            //     "width: 100%;" +
            //     "height: 100%;" +
            //     // "z-index: 1;" +
            //     "position: absolute;" +
            //     "top: 0;" +
            //     "left: 0;";

            // let cssName = "slide-background-" + index + "::before," + " .slide-background-" + index + "::after";
            // styleTable[styleText_before_after] = {
            //     "name": cssName,
            //     "text": styleText_before_after
            // };


            // let styleText_after = "background-color: #" + clr_ary[1] + ";" +
            //     "mix-blend-mode: darken;";

            // cssName = "slide-background-" + index + "::after";
            // styleTable[styleText_after] = {
            //     "name": cssName,
            //     "text": styleText_after
            // };

            // let styleText_before = "background-color: #" + clr_ary[0] + ";" +
            //     "mix-blend-mode: lighten;";

            // cssName = "slide-background-" + index + "::before";
            // styleTable[styleText_before] = {
            //     "name": cssName,
            //     "text": styleText_before
            // };

        }
        //a:alphaModFix
        let aphaModFixNode = getTextByPathList(aBlipNode, ["a:alphaModFix", "attrs"])
        let imgOpacity = "";
        if (aphaModFixNode !== undefined && aphaModFixNode["amt"] !== undefined && aphaModFixNode["amt"] != "") {
            let amt = parseInt(aphaModFixNode["amt"]) / 100000;
            //var opacity = amt;
            imgOpacity = "opacity:" + amt + ";";

        }
        //a:tile

        let tileNode = getTextByPathList(bgPr, ["a:blipFill", "a:tile", "attrs"])
        let prop_style = "";
        if (tileNode !== undefined && tileNode["sx"] !== undefined) {
            let sx = (parseInt(tileNode["sx"]) / 100000);
            let sy = (parseInt(tileNode["sy"]) / 100000);
            let tx = (parseInt(tileNode["tx"]) / 100000);
            let ty = (parseInt(tileNode["ty"]) / 100000);
            let algn = tileNode["algn"]; //tl(top left),t(top), tr(top right), l(left), ctr(center), r(right), bl(bottom left), b(bottm) , br(bottom right)
            let flip = tileNode["flip"]; //none,x,y ,xy

            prop_style += "background-repeat: round;"; //repeat|repeat-x|repeat-y|no-repeat|space|round|initial|inherit;
            //prop_style += "background-size: 300px 100px;"; size (w,h,sx, sy) -TODO
            //prop_style += "background-position: 50% 40%;"; //offset (tx, ty) -TODO
        }
        //a:srcRect
        //a:stretch => a:fillRect =>attrs (l:-17000, r:-17000)
        let stretch = getTextByPathList(bgPr, ["a:blipFill", "a:stretch"]);
        if (stretch !== undefined) {
            let fillRect = getTextByPathList(stretch, ["a:fillRect"]);
            //console.log("getBgPicFill=>bgPr: ", bgPr)
            // let top = fillRect["t"], right = fillRect["r"], bottom = fillRect["b"], left = fillRect["l"];
            prop_style += "background-repeat: no-repeat;";
            prop_style += "background-position: center;";
            if (fillRect !== undefined) {
                //prop_style += "background-size: contain, cover;";
                prop_style += "background-size:  100% 100%;;";
            }
        }
        bgcolor = "background: url(" + picFillBase64 + ");  z-index: " + ordr + ";" + prop_style + imgOpacity;

        return bgcolor;
    }
    // function hexToRgbNew(hex) {
    //     let arrBuff = new ArrayBuffer(4);
    //     let vw = new DataView(arrBuff);
    //     vw.setUint32(0, parseInt(hex, 16), false);
    //     let arrByte = new Uint8Array(arrBuff);
    //     return arrByte[1] + "," + arrByte[2] + "," + arrByte[3];
    // }
    private async getShapeFill(node: any, pNode: any, isSvgMode: boolean, source: string | undefined): Promise<any> {

        // 1. presentationML
        // p:spPr/ [a:noFill, solidFill, gradFill, blipFill, pattFill, grpFill]
        // From slide
        //Fill Type:
        //console.log("getShapeFill ShapeFill: ", node, ", isSvgMode; ", isSvgMode)
        let fillType = this.getFillType(getTextByPathList(node, ["p:spPr"]));
        //var noFill = getTextByPathList(node, ["p:spPr", "a:noFill"]);
        let fillColor;
        if (fillType == "NO_FILL") {
            return isSvgMode ? "none" : "";//"background-color: initial;";
        } else if (fillType == "SOLID_FILL") {
            let shpFill = node["p:spPr"]["a:solidFill"];
            fillColor = this.getSolidFill(shpFill);
        } else if (fillType == "GRADIENT_FILL") {
            let shpFill = node["p:spPr"]["a:gradFill"];
            fillColor = this.getGradientFill(shpFill);
            //console.log("shpFill",shpFill,grndColor.color)
        } else if (fillType == "PATTERN_FILL") {
            let shpFill = node["p:spPr"]["a:pattFill"];
            fillColor = this.getPatternFill(shpFill);
        } else if (fillType == "PIC_FILL") {
            let shpFill = node["p:spPr"]["a:blipFill"];
            fillColor = await this.getPicFill(source, shpFill);
        }
        //console.log("getShapeFill ShapeFill: ", node, ", isSvgMode; ", isSvgMode, ", fillType: ", fillType, ", fillColor: ", fillColor, ", source: ", source)


        // 2. drawingML namespace
        if (fillColor === undefined) {
            let clrName = getTextByPathList(node, ["p:style", "a:fillRef"]);
            let idx = parseInt(getTextByPathList(node, ["p:style", "a:fillRef", "attrs", "idx"]));
            if (idx == 0 || idx == 1000) {
                //no fill
                return isSvgMode ? "none" : "";
            } else if (idx > 0 && idx < 1000) {
                // <a:fillStyleLst> fill
            } else if (idx > 1000) {
                //<a:bgFillStyleLst>
            }
            fillColor = this.getSolidFill(clrName);
        }
        // 3. is group fill
        if (fillColor === undefined) {
            let grpFill = getTextByPathList(node, ["p:spPr", "a:grpFill"]);
            if (grpFill !== undefined) {
                //fillColor = this.getSolidFill(clrName, undefined, undefined, undefined, this.resources);
                //get parent fill style - TODO
                //console.log("ShapeFill: grpFill: ", grpFill, ", pNode: ", pNode)
                let grpShpFill = pNode["p:grpSpPr"];
                let spShpNode = { "p:spPr": grpShpFill }
                return await this.getShapeFill(spShpNode, node, isSvgMode, source);
            } else if (fillType == "NO_FILL") {
                return isSvgMode ? "none" : "";
            }
        }
        //console.log("ShapeFill: fillColor: ", fillColor, ", fillType; ", fillType)

        if (fillColor !== undefined) {
            if (fillType == "GRADIENT_FILL") {
                if (isSvgMode) {
                    // console.log("GRADIENT_FILL color", fillColor.color[0])
                    return fillColor;
                } else {
                    let colorAry = fillColor.color;
                    let rot = fillColor.rot;

                    let bgcolor = "background: linear-gradient(" + rot + "deg,";
                    for (var i = 0; i < colorAry.length; i++) {
                        if (i == colorAry.length - 1) {
                            bgcolor += "#" + colorAry[i] + ");";
                        } else {
                            bgcolor += "#" + colorAry[i] + ", ";
                        }

                    }
                    return bgcolor;
                }
            } else if (fillType == "PIC_FILL") {
                if (isSvgMode) {
                    return fillColor;
                } else {

                    return "background-image:url(" + fillColor + ");";
                }
            } else if (fillType == "PATTERN_FILL") {
                /////////////////////////////////////////////////////////////Need to check -----------TODO
                // if (isSvgMode) {
                //     let color = tinycolor(fillColor);
                //     fillColor = color.toRgbString();

                //     return fillColor;
                // } else {
                let bgPtrn = "", bgSize = "", bgPos = "";
                bgPtrn = fillColor[0];
                if (fillColor[1] !== null && fillColor[1] !== undefined && fillColor[1] != "") {
                    bgSize = " background-size:" + fillColor[1] + ";";
                }
                if (fillColor[2] !== null && fillColor[2] !== undefined && fillColor[2] != "") {
                    bgPos = " background-position:" + fillColor[2] + ";";
                }
                return "background: " + bgPtrn + ";" + bgSize + bgPos;
                //}
            } else {
                if (isSvgMode) {
                    let color = tinycolor(fillColor);
                    fillColor = color.toRgbString();

                    return fillColor;
                } else {
                    //console.log(node,"fillColor: ",fillColor,"fillType: ",fillType,"isSvgMode: ",isSvgMode)
                    return "background-color: #" + fillColor + ";";
                }
            }
        } else {
            if (isSvgMode) {
                return "none";
            } else {
                return "background-color: inherit;";
            }

        }

    }
    ///////////////////////Amir//////////////////////////////
    private getFillType(node: any) {
        //Need to test/////////////////////////////////////////////
        //SOLID_FILL
        //PIC_FILL
        //GRADIENT_FILL
        //PATTERN_FILL
        //NO_FILL
        let fillType = "";
        if (node["a:noFill"] !== undefined) {
            fillType = "NO_FILL";
        }
        if (node["a:solidFill"] !== undefined) {
            fillType = "SOLID_FILL";
        }
        if (node["a:gradFill"] !== undefined) {
            fillType = "GRADIENT_FILL";
        }
        if (node["a:pattFill"] !== undefined) {
            fillType = "PATTERN_FILL";
        }
        if (node["a:blipFill"] !== undefined) {
            fillType = "PIC_FILL";
        }
        if (node["a:grpFill"] !== undefined) {
            fillType = "GROUP_FILL";
        }


        return fillType;
    }
    private getGradientFill(node: any) {
        //console.log("getGradientFill: node", node)
        let gsLst = node["a:gsLst"]["a:gs"];
        //get start color
        let color_ary = [];
        let tint_ary = [];
        for (var i = 0; i < gsLst.length; i++) {
            let lo_tint;
            let lo_color = this.getSolidFill(gsLst[i]);
            //console.log("lo_color",lo_color)
            color_ary[i] = lo_color;
        }
        //get rot
        let lin = node["a:lin"];
        let rot = 0;
        if (lin !== undefined) {
            rot = angleToDegrees(lin["attrs"]?.["ang"]) + 90;
        }
        return {
            "color": color_ary,
            "rot": rot
        }
    }
    private async getPicFill(type: string | undefined, node: any) {
        //Need to test/////////////////////////////////////////////
        //rId
        //TODO - Image Properties - Tile, Stretch, or Display Portion of Image
        //(http://officeopenxml.com/drwPic-tile.php)
        let img;
        let rId = node["a:blip"]["attrs"]?.["r:embed"];
        let imgPath;
        //console.log("getPicFill(...) rId: ", rId, ", this.resources: ", this.resources, ", type: ", type)
        if (type == "slideBg" || type == "slide") {
            imgPath = getTextByPathList(this.resources, ["slideResObj", rId, "target"]);
        } else if (type == "slideLayoutBg") {
            imgPath = getTextByPathList(this.resources, ["layoutResObj", rId, "target"]);
        } else if (type == "slideMasterBg") {
            imgPath = getTextByPathList(this.resources, ["masterResObj", rId, "target"]);
        } else if (type == "themeBg") {
            imgPath = getTextByPathList(this.resources, ["themeResObj", rId, "target"]);
        } else if (type == "diagramBg") {
            imgPath = getTextByPathList(this.resources, ["diagramResObj", rId, "target"]);
        }
        if (imgPath === undefined) {
            return undefined;
        }
        img = getTextByPathList(this.resources, ["loaded-images", imgPath]); //, type, rId
        if (img === undefined) {
            imgPath = escapeHtml(imgPath);


            img = await this.options.resourceReader?.getImgBase64Data(imgPath);
            //this.resources["loaded-images"][imgPath] = img; //"defaultTextStyle": defaultTextStyle,
            if (img) setTextByPathList(this.resources, ["loaded-images", imgPath], img); //, type, rId
        }
        return img;
    }
    private getPatternFill(node: any) {
        //https://developer.mozilla.org/en-US/docs/Web/CSS/CSS_Images/Using_CSS_gradients
        //https://cssgradient.io/blog/css-gradient-text/
        //https://css-tricks.com/background-patterns-simplified-by-conic-gradients/
        //https://stackoverflow.com/questions/6705250/how-to-get-a-pattern-into-a-written-text-via-css
        //https://stackoverflow.com/questions/14072142/striped-text-in-css
        //https://css-tricks.com/stripes-css/
        //https://yuanchuan.dev/gradient-shapes/
        let fgColor, bgColor, prst = "";
        let bgClr = node["a:bgClr"];
        let fgClr = node["a:fgClr"];
        prst = node["attrs"]?.["prst"];
        fgColor = this.getSolidFill(fgClr);
        bgColor = this.getSolidFill(bgClr);
        //var angl_ary = getAnglefromParst(prst);
        //var ptrClr = "repeating-linear-gradient(" + angl + "deg,  #" + bgColor + ",#" + fgColor + " 2px);"
        //linear-gradient(0deg, black 10 %, transparent 10 %, transparent 90 %, black 90 %, black), 
        //linear-gradient(90deg, black 10 %, transparent 10 %, transparent 90 %, black 90 %, black);
        let linear_gradient = getLinerGrandient(prst, bgColor, fgColor);
        //console.log("getPatternFill: node:", node, ", prst: ", prst, ", fgColor: ", fgColor, ", bgColor:", bgColor, ', linear_gradient: ', linear_gradient)
        return linear_gradient;
    }

    private getSolidFill(node: any, clrMap?: any, phClr?: string) {

        if (node === undefined) {
            return undefined;
        }

        //console.log("getSolidFill node: ", node)
        let color: string | undefined = "";
        let clrNode;
        if (node["a:srgbClr"] !== undefined) {
            clrNode = node["a:srgbClr"];
            color = getTextByPathList(clrNode, ["attrs", "val"]); //#...
        } else if (node["a:schemeClr"] !== undefined) { //a:schemeClr
            clrNode = node["a:schemeClr"];
            let schemeClr = getTextByPathList(clrNode, ["attrs", "val"]);
            color = this.getSchemeColorFromTheme("a:" + schemeClr, clrMap, phClr);
            //console.log("schemeClr: ", schemeClr, "color: ", color)
        } else if (node["a:scrgbClr"] !== undefined) {
            clrNode = node["a:scrgbClr"];
            //<a:scrgbClr r="50%" g="50%" b="50%"/>  //Need to test/////////////////////////////////////////////
            let defBultColorVals = clrNode["attrs"] ?? {};
            let red = (defBultColorVals["r"].indexOf("%") != -1) ? defBultColorVals["r"].split("%").shift() : defBultColorVals["r"];
            let green = (defBultColorVals["g"].indexOf("%") != -1) ? defBultColorVals["g"].split("%").shift() : defBultColorVals["g"];
            let blue = (defBultColorVals["b"].indexOf("%") != -1) ? defBultColorVals["b"].split("%").shift() : defBultColorVals["b"];
            //var scrgbClr = red + "," + green + "," + blue;
            color = toHex(255 * (Number(red) / 100)) + toHex(255 * (Number(green) / 100)) + toHex(255 * (Number(blue) / 100));
            //console.log("scrgbClr: " + scrgbClr);

        } else if (node["a:prstClr"] !== undefined) {
            clrNode = node["a:prstClr"];
            //<a:prstClr val="black"/>  //Need to test/////////////////////////////////////////////
            let prstClr = getTextByPathList(clrNode, ["attrs", "val"]); //node["a:prstClr"]["attrs"]?.["val"];
            color = colorMap.get(prstClr);
            //console.log("blip prstClr: ", prstClr, " => hexClr: ", color);
        } else if (node["a:hslClr"] !== undefined) {
            clrNode = node["a:hslClr"];
            //<a:hslClr hue="14400000" sat="100%" lum="50%"/>  //Need to test/////////////////////////////////////////////
            let defBultColorVals = clrNode["attrs"] ?? {};
            let hue = Number(defBultColorVals["hue"]) / 100000;
            let sat = Number((defBultColorVals["sat"].indexOf("%") != -1) ? defBultColorVals["sat"].split("%").shift() : defBultColorVals["sat"]) / 100;
            let lum = Number((defBultColorVals["lum"].indexOf("%") != -1) ? defBultColorVals["lum"].split("%").shift() : defBultColorVals["lum"]) / 100;
            //var hslClr = defBultColorVals["hue"] + "," + defBultColorVals["sat"] + "," + defBultColorVals["lum"];
            let hsl2rgb = hslToRgb(hue, sat, lum);
            color = toHex(hsl2rgb.r) + toHex(hsl2rgb.g) + toHex(hsl2rgb.b);
            //defBultColor = cnvrtHslColor2Hex(hslClr); //TODO
            // console.log("hslClr: " + hslClr);
        } else if (node["a:sysClr"] !== undefined) {
            clrNode = node["a:sysClr"];
            //<a:sysClr val="windowText" lastClr="000000"/>  //Need to test/////////////////////////////////////////////
            let sysClr = getTextByPathList(clrNode, ["attrs", "lastClr"]);
            if (sysClr !== undefined) {
                color = sysClr;
            }
        }
        //console.log("color: [%cstart]", "color: #" + color, tinycolor(color).toHslString(), color)

        //fix color -------------------------------------------------------- TODO 
        //
        //1. "alpha":
        //Specifies the opacity as expressed by a percentage value.
        // [Example: The following represents a green solid fill which is 50 % opaque
        // < a: solidFill >
        //     <a:srgbClr val="00FF00">
        //         <a:alpha val="50%" />
        //     </a:srgbClr>
        // </a: solidFill >
        let isAlpha = false;
        let alpha = parseInt(getTextByPathList(clrNode, ["a:alpha", "attrs", "val"])) / 100000;
        //console.log("alpha: ", alpha)
        if (!isNaN(alpha)) {
            // let al_color = new colz.Color(color);
            // al_color.setAlpha(alpha);
            // let ne_color = al_color.rgba.toString();
            // color = (rgba2hex(ne_color))
            let al_color = tinycolor(color!);
            al_color.setAlpha(alpha);
            color = al_color.toHex8()
            isAlpha = true;
            //console.log("al_color: ", al_color, ", color: ", color)
        }
        //2. "alphaMod":
        // Specifies the opacity as expressed by a percentage relative to the input color.
        //     [Example: The following represents a green solid fill which is 50 % opaque
        //     < a: solidFill >
        //         <a:srgbClr val="00FF00">
        //             <a:alphaMod val="50%" />
        //         </a:srgbClr>
        //     </a: solidFill >
        //3. "alphaOff":
        // Specifies the opacity as expressed by a percentage offset increase or decrease to the
        // input color.Increases never increase the opacity beyond 100 %, decreases never decrease
        // the opacity below 0 %.
        // [Example: The following represents a green solid fill which is 90 % opaque
        //     < a: solidFill >
        //         <a:srgbClr val="00FF00">
        //             <a:alphaOff val="-10%" />
        //         </a:srgbClr>
        //     </a: solidFill >

        //4. "blue":
        //Specifies the value of the blue component.The assigned value is specified as a
        //percentage with 0 % indicating minimal blue and 100 % indicating maximum blue.
        //  [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
        //      to value RRGGBB = (00, FF, FF)
        //          <a: solidFill >
        //              <a:srgbClr val="00FF00">
        //                  <a:blue val="100%" />
        //              </a:srgbClr>
        //          </a: solidFill >
        //5. "blueMod"
        // Specifies the blue component as expressed by a percentage relative to the input color
        // component.Increases never increase the blue component beyond 100 %, decreases
        // never decrease the blue component below 0 %.
        // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, 00, FF)
        //     to value RRGGBB = (00, 00, 80)
        //     < a: solidFill >
        //         <a:srgbClr val="0000FF">
        //             <a:blueMod val="50%" />
        //         </a:srgbClr>
        //     </a: solidFill >
        //6. "blueOff"
        // Specifies the blue component as expressed by a percentage offset increase or decrease
        // to the input color component.Increases never increase the blue component
        // beyond 100 %, decreases never decrease the blue component below 0 %.
        // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, 00, FF)
        // to value RRGGBB = (00, 00, CC)
        //     < a: solidFill >
        //         <a:srgbClr val="00FF00">
        //             <a:blueOff val="-20%" />
        //         </a:srgbClr>
        //     </a: solidFill >

        //7. "comp" - This element specifies that the color rendered should be the complement of its input color with the complement
        // being defined as such.Two colors are called complementary if, when mixed they produce a shade of grey.For
        // instance, the complement of red which is RGB(255, 0, 0) is cyan.(<a:comp/>)

        //8. "gamma" - This element specifies that the output color rendered by the generating application should be the sRGB gamma
        //              shift of the input color.

        //9. "gray" - This element specifies a grayscale of its input color, taking into relative intensities of the red, green, and blue
        //              primaries.

        //10. "green":
        // Specifies the value of the green component. The assigned value is specified as a
        // percentage with 0 % indicating minimal green and 100 % indicating maximum green.
        // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, 00, FF)
        // to value RRGGBB = (00, FF, FF)
        //     < a: solidFill >
        //         <a:srgbClr val="0000FF">
        //             <a:green val="100%" />
        //         </a:srgbClr>
        //     </a: solidFill >
        //11. "greenMod":
        // Specifies the green component as expressed by a percentage relative to the input color
        // component.Increases never increase the green component beyond 100 %, decreases
        // never decrease the green component below 0 %.
        // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
        // to value RRGGBB = (00, 80, 00)
        //     < a: solidFill >
        //         <a:srgbClr val="00FF00">
        //             <a:greenMod val="50%" />
        //         </a:srgbClr>
        //     </a: solidFill >
        //12. "greenOff":
        // Specifies the green component as expressed by a percentage offset increase or decrease
        // to the input color component.Increases never increase the green component
        // beyond 100 %, decreases never decrease the green component below 0 %.
        // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
        // to value RRGGBB = (00, CC, 00)
        //     < a: solidFill >
        //         <a:srgbClr val="00FF00">
        //             <a:greenOff val="-20%" />
        //         </a:srgbClr>
        //     </a: solidFill >

        //13. "hue" (This element specifies a color using the HSL color model):
        // This element specifies the input color with the specified hue, but with its saturation and luminance unchanged.
        // < a: solidFill >
        //     <a:hslClr hue="14400000" sat="100%" lum="50%">
        // </a:solidFill>
        // <a:solidFill>
        //     <a:hslClr hue="0" sat="100%" lum="50%">
        //         <a:hue val="14400000"/>
        //     <a:hslClr/>
        // </a:solidFill>

        //14. "hueMod" (This element specifies a color using the HSL color model):
        // Specifies the hue as expressed by a percentage relative to the input color.
        // [Example: The following manipulates the fill color from having RGB value RRGGBB = (00, FF, 00) to value RRGGBB = (FF, FF, 00)
        //         < a: solidFill >
        //             <a:srgbClr val="00FF00">
        //                 <a:hueMod val="50%" />
        //             </a:srgbClr>
        //         </a: solidFill >

        let hueMod = parseInt(getTextByPathList(clrNode, ["a:hueMod", "attrs", "val"])) / 100000;
        //console.log("hueMod: ", hueMod)
        if (!isNaN(hueMod)) {
            color = applyHueMod(color!, hueMod, isAlpha);
        }
        //15. "hueOff"(This element specifies a color using the HSL color model):
        // Specifies the actual angular value of the shift.The result of the shift shall be between 0
        // and 360 degrees.Shifts resulting in angular values less than 0 are treated as 0. Shifts
        // resulting in angular values greater than 360 are treated as 360.
        // [Example:
        //     The following increases the hue angular value by 10 degrees.
        //     < a: solidFill >
        //         <a:hslClr hue="0" sat="100%" lum="50%"/>
        //             <a:hueOff val="600000"/>
        //     </a: solidFill >
        //var hueOff = parseInt(getTextByPathList(clrNode, ["a:hueOff", "attrs", "val"])) / 100000;
        // if (!isNaN(hueOff)) {
        //     //console.log("hueOff: ", hueOff, " (TODO)")
        //     //color = applyHueOff(color, hueOff, isAlpha);
        // }

        //16. "inv" (inverse)
        //This element specifies the inverse of its input color.
        //The inverse of red (1, 0, 0) is cyan (0, 1, 1 ).
        // The following represents cyan, the inverse of red:
        // <a:solidFill>
        //     <a:srgbClr val="FF0000">
        //         <a:inv />
        //     </a:srgbClr>
        // </a:solidFill>

        //17. "invGamma" - This element specifies that the output color rendered by the generating application should be the inverse sRGB
        //                  gamma shift of the input color.

        //18. "lum":
        // This element specifies the input color with the specified luminance, but with its hue and saturation unchanged.
        // Typically luminance values fall in the range[0 %, 100 %].
        // The following two solid fills are equivalent:
        // <a:solidFill>
        //     <a:hslClr hue="14400000" sat="100%" lum="50%">
        // </a:solidFill>
        // <a:solidFill>
        //     <a:hslClr hue="14400000" sat="100%" lum="0%">
        //         <a:lum val="50%" />
        //     <a:hslClr />
        // </a:solidFill>
        // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
        // to value RRGGBB = (00, 66, 00)
        //     < a: solidFill >
        //         <a:srgbClr val="00FF00">
        //             <a:lum val="20%" />
        //         </a:srgbClr>
        //     </a: solidFill >
        // end example]
        //19. "lumMod":
        // Specifies the luminance as expressed by a percentage relative to the input color.
        // Increases never increase the luminance beyond 100 %, decreases never decrease the
        // luminance below 0 %.
        // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
        //     to value RRGGBB = (00, 75, 00)
        //     < a: solidFill >
        //         <a:srgbClr val="00FF00">
        //             <a:lumMod val="50%" />
        //         </a:srgbClr>
        //     </a: solidFill >
        // end example]
        let lumMod = parseInt(getTextByPathList(clrNode, ["a:lumMod", "attrs", "val"])) / 100000;
        //console.log("lumMod: ", lumMod)
        if (!isNaN(lumMod)) {
            color = applyLumMod(color!, lumMod, isAlpha);
        }
        //var lumMod_color = applyLumMod(color, 0.5);
        //console.log("lumMod_color: ", lumMod_color)
        //20. "lumOff"
        // Specifies the luminance as expressed by a percentage offset increase or decrease to the
        // input color.Increases never increase the luminance beyond 100 %, decreases never
        // decrease the luminance below 0 %.
        // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
        //     to value RRGGBB = (00, 99, 00)
        //     < a: solidFill >
        //         <a:srgbClr val="00FF00">
        //             <a:lumOff val="-20%" />
        //         </a:srgbClr>
        //     </a: solidFill >
        let lumOff = parseInt(getTextByPathList(clrNode, ["a:lumOff", "attrs", "val"])) / 100000;
        //console.log("lumOff: ", lumOff)
        if (!isNaN(lumOff)) {
            color = applyLumOff(color!, lumOff, isAlpha);
        }


        //21. "red":
        // Specifies the value of the red component.The assigned value is specified as a percentage
        // with 0 % indicating minimal red and 100 % indicating maximum red.
        // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
        //     to value RRGGBB = (FF, FF, 00)
        //     < a: solidFill >
        //         <a:srgbClr val="00FF00">
        //             <a:red val="100%" />
        //         </a:srgbClr>
        //     </a: solidFill >
        //22. "redMod":
        // Specifies the red component as expressed by a percentage relative to the input color
        // component.Increases never increase the red component beyond 100 %, decreases never
        // decrease the red component below 0 %.
        // [Example: The following manipulates the fill from having RGB value RRGGBB = (FF, 00, 00)
        //     to value RRGGBB = (80, 00, 00)
        //     < a: solidFill >
        //         <a:srgbClr val="FF0000">
        //             <a:redMod val="50%" />
        //         </a:srgbClr>
        //     </a: solidFill >
        //23. "redOff":
        // Specifies the red component as expressed by a percentage offset increase or decrease to
        // the input color component.Increases never increase the red component beyond 100 %,
        //     decreases never decrease the red component below 0 %.
        //     [Example: The following manipulates the fill from having RGB value RRGGBB = (FF, 00, 00)
        //     to value RRGGBB = (CC, 00, 00)
        //     < a: solidFill >
        //         <a:srgbClr val="FF0000">
        //             <a:redOff val="-20%" />
        //         </a:srgbClr>
        //     </a: solidFill >

        //23. "sat":
        // This element specifies the input color with the specified saturation, but with its hue and luminance unchanged.
        // Typically saturation values fall in the range[0 %, 100 %].
        // [Example:
        //     The following two solid fills are equivalent:
        //     <a:solidFill>
        //         <a:hslClr hue="14400000" sat="100%" lum="50%">
        //     </a:solidFill>
        //     <a:solidFill>
        //         <a:hslClr hue="14400000" sat="0%" lum="50%">
        //             <a:sat val="100000" />
        //         <a:hslClr />
        //     </a:solidFill>
        // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
        //     to value RRGGBB = (40, C0, 40)
        //     < a: solidFill >
        //         <a:srgbClr val="00FF00">
        //             <a:sat val="50%" />
        //         </a:srgbClr>
        //     <a: solidFill >
        // end example]

        //24. "satMod":
        // Specifies the saturation as expressed by a percentage relative to the input color.
        // Increases never increase the saturation beyond 100 %, decreases never decrease the
        // saturation below 0 %.
        // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
        //     to value RRGGBB = (66, 99, 66)
        //     < a: solidFill >
        //         <a:srgbClr val="00FF00">
        //             <a:satMod val="20%" />
        //         </a:srgbClr>
        //     </a: solidFill >
        let satMod = parseInt(getTextByPathList(clrNode, ["a:satMod", "attrs", "val"])) / 100000;
        if (!isNaN(satMod)) {
            color = applySatMod(color!, satMod, isAlpha);
        }
        //25. "satOff":
        // Specifies the saturation as expressed by a percentage offset increase or decrease to the
        // input color.Increases never increase the saturation beyond 100 %, decreases never
        // decrease the saturation below 0 %.
        // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
        //     to value RRGGBB = (19, E5, 19)
        //     < a: solidFill >
        //         <a:srgbClr val="00FF00">
        //             <a:satOff val="-20%" />
        //         </a:srgbClr>
        //     </a: solidFill >
        // let satOff = parseInt(getTextByPathList(clrNode, ["a:satOff", "attrs", "val"])) / 100000;
        // if (!isNaN(satOff)) {
        //     console.log("satOff: ", satOff, " (TODO)")
        // }

        //26. "shade":
        // This element specifies a darker version of its input color.A 10 % shade is 10 % of the input color combined with 90 % black.
        // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
        //     to value RRGGBB = (00, BC, 00)
        //     < a: solidFill >
        //         <a:srgbClr val="00FF00">
        //             <a:shade val="50%" />
        //         </a:srgbClr>
        //     </a: solidFill >
        // end example]
        let shade = parseInt(getTextByPathList(clrNode, ["a:shade", "attrs", "val"])) / 100000;
        if (!isNaN(shade)) {
            color = applyShade(color!, shade, isAlpha);
        }
        //27.  "tint":
        // This element specifies a lighter version of its input color.A 10 % tint is 10 % of the input color combined with
        // 90 % white.
        // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
        //     to value RRGGBB = (BC, FF, BC)
        //     < a: solidFill >
        //         <a:srgbClr val="00FF00">
        //             <a:tint val="50%" />
        //         </a:srgbClr>
        //     </a: solidFill >
        let tint = parseInt(getTextByPathList(clrNode, ["a:tint", "attrs", "val"])) / 100000;
        if (!isNaN(tint)) {
            color = applyTint(color!, tint, isAlpha);
        }
        //console.log("color [%cfinal]: ", "color: #" + color, tinycolor(color).toHslString(), color)

        return color;
    }
    private getSchemeColorFromTheme(schemeClr: string, clrMap?: any, phClr?: string) {
        //<p:clrMap ...> in slide master
        // e.g. tx2="dk2" bg2="lt2" tx1="dk1" bg1="lt1" slideLayoutClrOvride
        //console.log("getSchemeColorFromTheme: schemeClr: ", schemeClr, ",clrMap: ", clrMap)
        let slideLayoutClrOvride;
        if (clrMap !== undefined) {
            slideLayoutClrOvride = clrMap;//getTextByPathList(clrMap, ["p:sldMaster", "p:clrMap", "attrs"])
        } else {
            let sldClrMapOvr = getTextByPathList(this.resources?.slideContent, ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
            if (sldClrMapOvr !== undefined) {
                slideLayoutClrOvride = sldClrMapOvr;
            } else {
                let sldClrMapOvr = getTextByPathList(this.resources?.slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                if (sldClrMapOvr !== undefined) {
                    slideLayoutClrOvride = sldClrMapOvr;
                } else {
                    slideLayoutClrOvride = getTextByPathList(this.resources?.slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
                }

            }
        }
        //console.log("getSchemeColorFromTheme slideLayoutClrOvride: ", slideLayoutClrOvride);
        let schmClrName = schemeClr.substr(2);
        let color;
        if (schmClrName == "phClr" && phClr !== undefined) {
            color = phClr;
        } else {
            if (slideLayoutClrOvride !== undefined) {
                switch (schmClrName) {
                    case "tx1":
                    case "tx2":
                    case "bg1":
                    case "bg2":
                        schemeClr = "a:" + slideLayoutClrOvride[schmClrName];
                        break;
                }
            } else {
                switch (schmClrName) {
                    case "tx1":
                        schemeClr = "a:dk1";
                        break;
                    case "tx2":
                        schemeClr = "a:dk2";
                        break;
                    case "bg1":
                        schemeClr = "a:lt1";
                        break;
                    case "bg2":
                        schemeClr = "a:lt2";
                        break;
                }
            }
            //console.log("getSchemeColorFromTheme:  schemeClr: ", schemeClr);
            let refNode = getTextByPathList(this.resources?.themeContent, ["a:theme", "a:themeElements", "a:clrScheme", schemeClr]);
            color = getTextByPathList(refNode, ["a:srgbClr", "attrs", "val"]);
            //console.log("themeContent: color", color);
            if (color === undefined && refNode !== undefined) {
                color = getTextByPathList(refNode, ["a:sysClr", "attrs", "lastClr"]);
            }
        }
        //console.log(color)
        return color;
    }

    private extractChartData(serNode: any) {

        let dataMat = new Array();

        if (serNode === undefined) {
            return dataMat;
        }

        if (serNode["c:xVal"] !== undefined) {
            let dataRow = new Array();
            eachElement(serNode["c:xVal"]["c:numRef"]["c:numCache"]["c:pt"], (innerNode: any) => {
                dataRow.push(parseFloat(innerNode["c:v"]));
                return "";
            });
            dataMat.push(dataRow);
            dataRow = new Array();
            eachElement(serNode["c:yVal"]["c:numRef"]["c:numCache"]["c:pt"], (innerNode: any) => {
                dataRow.push(parseFloat(innerNode["c:v"]));
                return "";
            });
            dataMat.push(dataRow);
        } else {
            eachElement(serNode, (innerNode: any, index: number) => {
                let dataRow = new Array();
                let colName = getTextByPathList(innerNode, ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"]) || index;

                // Category (string or number)
                let rowNames: any = {};
                if (getTextByPathList(innerNode, ["c:cat", "c:strRef", "c:strCache", "c:pt"]) !== undefined) {
                    eachElement(innerNode["c:cat"]["c:strRef"]["c:strCache"]["c:pt"], (innerNode: any) => {
                        rowNames[innerNode["attrs"]?.["idx"]] = innerNode["c:v"];
                        return "";
                    });
                } else if (getTextByPathList(innerNode, ["c:cat", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
                    eachElement(innerNode["c:cat"]["c:numRef"]["c:numCache"]["c:pt"], (innerNode: any) => {
                        rowNames[innerNode["attrs"]?.["idx"]] = innerNode["c:v"];
                        return "";
                    });
                }

                // Value
                if (getTextByPathList(innerNode, ["c:val", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
                    eachElement(innerNode["c:val"]["c:numRef"]["c:numCache"]["c:pt"], (innerNode: any) => {
                        dataRow.push({ x: innerNode["attrs"]?.["idx"], y: parseFloat(innerNode["c:v"]) });
                        return "";
                    });
                }

                dataMat.push({ key: colName, values: dataRow, xlabels: rowNames });
                return "";
            });
        }

        return dataMat;
    }

    private getSvgImagePattern(node: any, fill: string, shpId: string) {
        let pic_dim = getBase64ImageDimensions(fill);
        let width = pic_dim?.[0] ?? 0;
        let height = pic_dim?.[1] ?? 0;
        //console.log("getSvgImagePattern node:", node);
        let blipFillNode = node["p:spPr"]["a:blipFill"];
        let tileNode = getTextByPathList(blipFillNode, ["a:tile", "attrs"])
        let sx, sy;
        if (tileNode !== undefined && tileNode["sx"] !== undefined) {
            sx = (parseInt(tileNode["sx"]) / 100000) * width;
            sy = (parseInt(tileNode["sy"]) / 100000) * height;
        }

        let blipNode = node["p:spPr"]["a:blipFill"]["a:blip"];
        let tialphaModFixNode = getTextByPathList(blipNode, ["a:alphaModFix", "attrs"])
        let imgOpacity = "";
        if (tialphaModFixNode !== undefined && tialphaModFixNode["amt"] !== undefined && tialphaModFixNode["amt"] != "") {
            let amt = parseInt(tialphaModFixNode["amt"]) / 100000;
            let opacity = amt;
            let imgOpacity = "opacity='" + opacity + "'";

        }
        let ptrn;
        if (sx !== undefined && sx != 0) {
            ptrn = '<pattern id="imgPtrn_' + shpId + '" x="0" y="0"  width="' + sx + '" height="' + sy + '" patternUnits="userSpaceOnUse">';
        } else {
            ptrn = '<pattern id="imgPtrn_' + shpId + '"  patternContentUnits="objectBoundingBox"  width="1" height="1">';
        }
        let duotoneNode = getTextByPathList(blipNode, ["a:duotone"])
        let fillterNode = "";
        let filterUrl = "";
        if (duotoneNode !== undefined) {
            //console.log("pic duotoneNode: ", duotoneNode)
            let clr_ary: any[] = [];
            Object.keys(duotoneNode).forEach((clr_type) => {
                //Object.keys(duotoneNode[clr_type]).forEach(function (clr) {
                //console.log("blip pic duotone clr: ", duotoneNode[clr_type][clr], clr)
                if (clr_type != "attrs") {
                    let obj: any = {};
                    obj[clr_type] = duotoneNode[clr_type];
                    //console.log("blip pic duotone obj: ", obj)
                    let hexClr = this.getSolidFill(obj)
                    //clr_ary.push();

                    let color = tinycolor("#" + hexClr);
                    clr_ary.push(color.toRgb()); // { r: 255, g: 0, b: 0, a: 1 }
                }
                // })
            })

            if (clr_ary.length == 2) {

                fillterNode = '<filter id="svg_image_duotone"> ' +
                    '<feColorMatrix type="matrix" values=".33 .33 .33 0 0' +
                    '.33 .33 .33 0 0' +
                    '.33 .33 .33 0 0' +
                    '0 0 0 1 0">' +
                    '</feColorMatrix>' +
                    '<feComponentTransfer color-interpolation-filters="sRGB">' +
                    //clr_ary.forEach(function(clr){
                    '<feFuncR type="table" tableValues="' + clr_ary[0].r / 255 + ' ' + clr_ary[1].r / 255 + '"></feFuncR>' +
                    '<feFuncG type="table" tableValues="' + clr_ary[0].g / 255 + ' ' + clr_ary[1].g / 255 + '"></feFuncG>' +
                    '<feFuncB type="table" tableValues="' + clr_ary[0].b / 255 + ' ' + clr_ary[1].b / 255 + '"></feFuncB>' +
                    //});
                    '</feComponentTransfer>' +
                    ' </filter>';
            }

            filterUrl = 'filter="url(#svg_image_duotone)"';

            ptrn += fillterNode;
        }

        fill = escapeHtml(fill);
        if (sx !== undefined && sx != 0) {
            ptrn += '<image  xlink:href="' + fill + '" x="0" y="0" width="' + sx + '" height="' + sy + '" ' + imgOpacity + ' ' + filterUrl + '></image>';
        } else {
            ptrn += '<image  xlink:href="' + fill + '" preserveAspectRatio="none" width="1" height="1" ' + imgOpacity + ' ' + filterUrl + '></image>';
        }
        ptrn += '</pattern>';

        //console.log("getSvgImagePattern(...) pic_dim:", pic_dim, ", fillColor: ", fill, ", blipNode: ", blipNode, ",sx: ", sx, ", sy: ", sy, ", clr_ary: ", clr_ary, ", ptrn: ", ptrn)

        return ptrn;
    }

    private processMsgQueue() {
        for (var i = 0; i < this.MsgQueue.length; i++) {
            this.processSingleMsg(this.MsgQueue[i].data);
        }
    }

    private processSingleMsg(d: any) {
        var chartID = d.chartID;
        var chartType = d.chartType;
        var chartData = d.chartData;

        var data = [];

        var chart = null;
        const nv = window.nv;
        const d3 = window.d3;
        switch (chartType) {
            case "lineChart":
                data = chartData;
                chart = nv.models.lineChart()
                    .useInteractiveGuideline(true);
                chart.xAxis.tickFormat(function (d) { return chartData[0].xlabels[d] || d; });
                break;
            case "barChart":
                data = chartData;
                chart = nv.models.multiBarChart();
                chart.xAxis.tickFormat(function (d) { return chartData[0].xlabels[d] || d; });
                break;
            case "pieChart":
            case "pie3DChart":
                if (chartData.length > 0) {
                    data = chartData[0].values;
                }
                chart = nv.models.pieChart();
                break;
            case "areaChart":
                data = chartData;
                chart = nv.models.stackedAreaChart()
                    .clipEdge(true)
                    .useInteractiveGuideline(true);
                chart.xAxis.tickFormat(function (d) { return chartData[0].xlabels[d] || d; });
                break;
            case "scatterChart":

                for (var i = 0; i < chartData.length; i++) {
                    var arr = [];
                    for (var j = 0; j < chartData[i].length; j++) {
                        arr.push({ x: j, y: chartData[i][j] });
                    }
                    data.push({ key: 'data' + (i + 1), values: arr });
                }

                //data = chartData;
                chart = nv.models.scatterChart()
                    .showDistX(true)
                    .showDistY(true)
                    .color(d3.scale.category10().range());
                chart.xAxis.axisLabel('X').tickFormat(d3.format('.02f'));
                chart.yAxis.axisLabel('Y').tickFormat(d3.format('.02f'));
                break;
            default:
        }

        if (chart !== null) {
            d3.select("#" + chartID)
                .append("svg")
                .datum(data)
                .transition().duration(500)
                .call(chart);

            nv.utils.windowResize(chart.update);
        }
    }

    // d3 and nv d3 are initialized after the PPTX is rendered
    // so we need to call this function after the PPTX is rendered
    postRender() {
        if (this.postRenderDone) return;
        this.processMsgQueue();
        this.postRenderDone = true;
    }
}
