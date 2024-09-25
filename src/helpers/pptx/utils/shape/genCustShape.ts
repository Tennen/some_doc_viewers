import { getTextByPathList } from "../object";
import { shapeArc } from "../svg";
import { Border } from "./type";

export const genCustShape = (
    moveToNode: any,
    lnToNodes: any,
    cubicBezToNodes: any,
    arcToNodes: any,
    closeNode: any,
    cX: number,
    cY: number,
    fill: string,
    border: Border,
) => {
    let result = "";
    let multiSapeAry = [];
    if (moveToNode.length > 0) {
        //a:moveTo
        Object.keys(moveToNode).forEach((key) => {
            let moveToPtNode = moveToNode[key]["a:pt"];
            if (moveToPtNode !== undefined) {
                if (!Array.isArray(moveToPtNode)) {
                    moveToPtNode = [moveToPtNode];
                }
                Object.keys(moveToPtNode).forEach(function (key2) {
                    let ptObj: any = {};
                    let moveToNoPt = moveToPtNode[key2];
                    let spX = moveToNoPt["attrs"]?.["x"];//parseInt(moveToNoPt["attrs", "x"]) * this.options.slideFactor;
                    let spY = moveToNoPt["attrs"]?.["y"];//parseInt(moveToNoPt["attrs", "y"]) * this.options.slideFactor;
                    let ptOrdr = moveToNoPt["attrs"]?.["order"];
                    ptObj.type = "movto";
                    ptObj.order = ptOrdr;
                    ptObj.x = spX;
                    ptObj.y = spY;
                    multiSapeAry.push(ptObj);
                    //console.log(key2, lnToNoPt);

                });
            }
        });
        //a:lnTo
        if (lnToNodes !== undefined) {
            Object.keys(lnToNodes).forEach((key) => {
                let lnToPtNode = lnToNodes[key]["a:pt"];
                if (lnToPtNode !== undefined) {
                    if (!Array.isArray(lnToPtNode)) {
                        lnToPtNode = [lnToPtNode];
                    }
                    Object.keys(lnToPtNode).forEach(function (key2) {
                        let ptObj: any = {};
                        let lnToNoPt = lnToPtNode[key2];
                        let ptX = lnToNoPt["attrs"]?.["x"];
                        let ptY = lnToNoPt["attrs"]?.["y"];
                        let ptOrdr = lnToNoPt["attrs"]?.["order"];
                        ptObj.type = "lnto";
                        ptObj.order = ptOrdr;
                        ptObj.x = ptX;
                        ptObj.y = ptY;
                        multiSapeAry.push(ptObj);
                        //console.log(key2, lnToNoPt);
                    });
                }
            });
        }
        //a:cubicBezTo
        if (cubicBezToNodes !== undefined) {

            let cubicBezToPtNodesAry: any[] = [];
            //console.log("cubicBezToNodes: ", cubicBezToNodes, ", is arry: ", Array.isArray(cubicBezToNodes))
            if (!Array.isArray(cubicBezToNodes)) {
                cubicBezToNodes = [cubicBezToNodes];
            }
            Object.keys(cubicBezToNodes).forEach((key) => {
                //console.log("cubicBezTo[" + key + "]:");
                cubicBezToPtNodesAry.push(cubicBezToNodes[key]["a:pt"]);
            });

            //console.log("cubicBezToNodes: ", cubicBezToPtNodesAry)
            cubicBezToPtNodesAry.forEach(function (key2) {
                //console.log("cubicBezToPtNodesAry: key2 : ", key2)
                let nodeObj: any = {};
                nodeObj.type = "cubicBezTo";
                nodeObj.order = key2[0]["attrs"]?.["order"];
                let pts_ary: any[] = [];
                key2.forEach(function (pt: any) {
                    let pt_obj = {
                        x: pt["attrs"]?.["x"],
                        y: pt["attrs"]?.["y"]
                    }
                    pts_ary.push(pt_obj)
                })
                nodeObj.cubBzPt = pts_ary;//key2;
                multiSapeAry.push(nodeObj);
            });
        }
        //a:arcTo
        if (arcToNodes !== undefined) {
            let arcToNodesAttrs = arcToNodes["attrs"] ?? {};
            let arcOrder = arcToNodesAttrs["order"];
            let hR = arcToNodesAttrs["hR"];
            let wR = arcToNodesAttrs["wR"];
            let stAng = arcToNodesAttrs["stAng"];
            let swAng = arcToNodesAttrs["swAng"];
            let shftX = 0;
            let shftY = 0;
            let arcToPtNode = getTextByPathList(arcToNodes, ["a:pt", "attrs"]);
            if (arcToPtNode !== undefined) {
                shftX = arcToPtNode["x"];
                shftY = arcToPtNode["y"];
                //console.log("shftX: ",shftX," shftY: ",shftY)
            }
            let ptObj: any = {};
            ptObj.type = "arcTo";
            ptObj.order = arcOrder;
            ptObj.hR = hR;
            ptObj.wR = wR;
            ptObj.stAng = stAng;
            ptObj.swAng = swAng;
            ptObj.shftX = shftX;
            ptObj.shftY = shftY;
            multiSapeAry.push(ptObj);

        }
        //a:quadBezTo - TODO

        //a:close
        if (closeNode !== undefined) {

            if (!Array.isArray(closeNode)) {
                closeNode = [closeNode];
            }
            // Object.keys(closeNode).forEach((key) => {
            //     //console.log("cubicBezTo[" + key + "]:");
            //     cubicBezToPtNodesAry.push(closeNode[key]["a:pt"]);
            // });
            Object.keys(closeNode).forEach((key) => {
                //console.log("custShapType >> closeNode: key: ", key);
                let clsAttrs = closeNode[key]["attrs"] ?? {};
                //var clsAttrs = closeNode["attrs"];
                let clsOrder = clsAttrs["order"];
                let ptObj: any = {};
                ptObj.type = "close";
                ptObj.order = clsOrder;
                multiSapeAry.push(ptObj);

            });

        }

        // console.log("custShapType >> multiSapeAry: ", multiSapeAry);

        multiSapeAry.sort(function (a, b) {
            return a.order - b.order;
        });

        //console.log("custShapType >>sorted  multiSapeAry: ");
        //console.log(multiSapeAry);
        let k = 0;
        let isClose = false;
        let d = "";
        while (k < multiSapeAry.length) {

            if (multiSapeAry[k].type == "movto") {
                //start point
                let spX = parseInt(multiSapeAry[k].x) * cX;//slideFactor;
                let spY = parseInt(multiSapeAry[k].y) * cY;//slideFactor;
                // if (d == "") {
                //     d = "M" + spX + "," + spY;
                // } else {
                //     //shape without close : then close the shape and start new path
                //     result += "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                //         "' stroke='" + ((border === undefined) ? "" : border.color) + "' stroke-width='" + ((border === undefined) ? "" : border.width) + "' stroke-dasharray='" + ((border === undefined) ? "" : border.strokeDasharray) + "' ";
                //     result += "/>";

                //     if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                //         result += "marker-start='url(#markerTriangle_" + shpId + ")' ";
                //     }
                //     if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                //         result += "marker-end='url(#markerTriangle_" + shpId + ")' ";
                //     }
                //     result += "/>";

                //     d = "M" + spX + "," + spY;
                //     isClose = true;
                // }

                d += " M" + spX + "," + spY;

            } else if (multiSapeAry[k].type == "lnto") {
                let Lx = parseInt(multiSapeAry[k].x) * cX;//slideFactor;
                let Ly = parseInt(multiSapeAry[k].y) * cY;//slideFactor;
                d += " L" + Lx + "," + Ly;

            } else if (multiSapeAry[k].type == "cubicBezTo") {
                let Cx1 = parseInt(multiSapeAry[k].cubBzPt[0].x) * cX;//slideFactor;
                let Cy1 = parseInt(multiSapeAry[k].cubBzPt[0].y) * cY;//slideFactor;
                let Cx2 = parseInt(multiSapeAry[k].cubBzPt[1].x) * cX;//slideFactor;
                let Cy2 = parseInt(multiSapeAry[k].cubBzPt[1].y) * cY;//slideFactor;
                let Cx3 = parseInt(multiSapeAry[k].cubBzPt[2].x) * cX;//slideFactor;
                let Cy3 = parseInt(multiSapeAry[k].cubBzPt[2].y) * cY;//slideFactor;
                d += " C" + Cx1 + "," + Cy1 + " " + Cx2 + "," + Cy2 + " " + Cx3 + "," + Cy3;
            } else if (multiSapeAry[k].type == "arcTo") {
                let hR = parseInt(multiSapeAry[k].hR) * cX;//slideFactor;
                let wR = parseInt(multiSapeAry[k].wR) * cY;//slideFactor;
                let stAng = parseInt(multiSapeAry[k].stAng) / 60000;
                let swAng = parseInt(multiSapeAry[k].swAng) / 60000;
                //var shftX = parseInt(multiSapeAry[k].shftX) * this.options.slideFactor;
                //var shftY = parseInt(multiSapeAry[k].shftY) * this.options.slideFactor;
                let endAng = stAng + swAng;

                d += shapeArc(wR, hR, wR, hR, stAng, endAng, false);
            } else if (multiSapeAry[k].type == "quadBezTo") {
                console.log("custShapType: quadBezTo - TODO")

            } else if (multiSapeAry[k].type == "close") {
                // result += "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                //     "' stroke='" + ((border === undefined) ? "" : border.color) + "' stroke-width='" + ((border === undefined) ? "" : border.width) + "' stroke-dasharray='" + ((border === undefined) ? "" : border.strokeDasharray) + "' ";
                // result += "/>";
                // d = "";
                // isClose = true;

                d += "z";
            }
            k++;
        }
        //if (!isClose) {
        //only one "moveTo" and no "close"
        result += "<path d='" + d + "' fill='" + fill +
            "' stroke='" + ((border === undefined) ? "" : border.color) + "' stroke-width='" + ((border === undefined) ? "" : border.width) + "' stroke-dasharray='" + ((border === undefined) ? "" : border.strokeDasharray) + "' ";
        result += "/>";
        //console.log(result);
        return result;
    }
}