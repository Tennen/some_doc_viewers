export const getMiddleStops = (s: number) => {
    let sArry = ['0%', '100%'];
    if (s == 0) {
        return sArry;
    } else {
        let i = s;
        while (i--) {
            let middleStop = 100 - ((100 / (s + 1)) * (i + 1)), // AM: Ex - For 3 middle stops, progression will be 25%, 50%, and 75%, plus 0% and 100% at the ends.
                middleStopString = middleStop + "%";
            sArry.splice(-1, 0, middleStopString);
        } // AM: add into stopsArray before 100%
    }
    return sArry
}
export const SVGangle = (deg: number, svgHeight: number, svgWidth: number) => {
    let w = svgWidth,
        h = svgHeight,
        ang = deg,
        o = 2,
        n = 2,
        wc = w / 2,
        hc = h / 2,
        tx1 = 2,
        ty1 = 2,
        tx2 = 2,
        ty2 = 2,
        k = (((ang % 360) + 360) % 360),
        j = (360 - k) * Math.PI / 180,
        i = Math.tan(j),
        l = hc - i * wc;

    if (k == 0) {
        tx1 = w,
            ty1 = hc,
            tx2 = 0,
            ty2 = hc
    } else if (k < 90) {
        n = w,
            o = 0
    } else if (k == 90) {
        tx1 = wc,
            ty1 = 0,
            tx2 = wc,
            ty2 = h
    } else if (k < 180) {
        n = 0,
            o = 0
    } else if (k == 180) {
        tx1 = 0,
            ty1 = hc,
            tx2 = w,
            ty2 = hc
    } else if (k < 270) {
        n = 0,
            o = h
    } else if (k == 270) {
        tx1 = wc,
            ty1 = h,
            tx2 = wc,
            ty2 = 0
    } else {
        n = w,
            o = h;
    }
    // AM: I could not quite figure out what m, n, and o are supposed to represent from the original code on visualcsstools.com.
    let m = o + (n / i),
        x1 = Math.round(tx2 / w * 100 * 100) / 100,
        y1 = Math.round(ty2 / h * 100 * 100) / 100,
        x2 = Math.round(tx1 / w * 100 * 100) / 100,
        y2 = Math.round(ty1 / h * 100 * 100) / 100;
    tx1 = tx1 == 2 ? i * (m - l) / (Math.pow(i, 2) + 1) : tx1;
    ty1 = ty1 == 2 ? i * tx1 + l : ty1;
    tx2 = tx2 == 2 ? w - tx1 : tx2;
    ty2 = ty2 == 2 ? h - ty1 : ty2;
    return [x1, y1, x2, y2];
}

export const getLinerGrandient = (prst: string, bgColor: string | undefined, fgColor: string | undefined) => {
    // dashDnDiag (Dashed Downward Diagonal)-V
    // dashHorz (Dashed Horizontal)-V
    // dashUpDiag(Dashed Upward DIagonal)-V
    // dashVert(Dashed Vertical)-V
    // diagBrick(Diagonal Brick)-V
    // divot(Divot)-VX
    // dkDnDiag(Dark Downward Diagonal)-V
    // dkHorz(Dark Horizontal)-V
    // dkUpDiag(Dark Upward Diagonal)-V
    // dkVert(Dark Vertical)-V
    // dotDmnd(Dotted Diamond)-VX
    // dotGrid(Dotted Grid)-V
    // horzBrick(Horizontal Brick)-V
    // lgCheck(Large Checker Board)-V
    // lgConfetti(Large Confetti)-V
    // lgGrid(Large Grid)-V
    // ltDnDiag(Light Downward Diagonal)-V
    // ltHorz(Light Horizontal)-V
    // ltUpDiag(Light Upward Diagonal)-V
    // ltVert(Light Vertical)-V
    // narHorz(Narrow Horizontal)-V
    // narVert(Narrow Vertical)-V
    // openDmnd(Open Diamond)-V
    // pct10(10 %)-V
    // pct20(20 %)-V
    // pct25(25 %)-V
    // pct30(30 %)-V
    // pct40(40 %)-V
    // pct5(5 %)-V
    // pct50(50 %)-V
    // pct60(60 %)-V
    // pct70(70 %)-V
    // pct75(75 %)-V
    // pct80(80 %)-V
    // pct90(90 %)-V
    // smCheck(Small Checker Board) -V
    // smConfetti(Small Confetti)-V
    // smGrid(Small Grid) -V
    // solidDmnd(Solid Diamond)-V
    // sphere(Sphere)-V
    // trellis(Trellis)-VX
    // wave(Wave)-V
    // wdDnDiag(Wide Downward Diagonal)-V
    // wdUpDiag(Wide Upward Diagonal)-V
    // weave(Weave)-V
    // zigZag(Zig Zag)-V
    // shingle(Shingle)-V
    // plaid(Plaid)-V
    // cross (Cross)
    // diagCross(Diagonal Cross)
    // dnDiag(Downward Diagonal)
    // horz(Horizontal)
    // upDiag(Upward Diagonal)
    // vert(Vertical)
    switch (prst) {
        case "smGrid":
            return ["linear-gradient(to right,  #" + fgColor + " -1px, transparent 1px ), " +
                "linear-gradient(to bottom,  #" + fgColor + " -1px, transparent 1px)  #" + bgColor + ";", "4px 4px"];
        case "dotGrid":
            return ["linear-gradient(to right,  #" + fgColor + " -1px, transparent 1px ), " +
                "linear-gradient(to bottom,  #" + fgColor + " -1px, transparent 1px)  #" + bgColor + ";", "8px 8px"];
        case "lgGrid":
            return ["linear-gradient(to right,  #" + fgColor + " -1px, transparent 1.5px ), " +
                "linear-gradient(to bottom,  #" + fgColor + " -1px, transparent 1.5px)  #" + bgColor + ";", "8px 8px"];
        case "wdUpDiag":
            //return ["repeating-linear-gradient(-45deg,  #" + bgColor + ", #" + bgColor + " 1px,#" + fgColor + " 5px);"];
            return ["repeating-linear-gradient(-45deg, transparent 1px , transparent 4px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
            // return ["linear-gradient(45deg, transparent 0%, transparent calc(50% - 1px),  #" + fgColor + " 50%, transparent calc(50% + 1px),  transparent 100%) " +
            //     "#" + bgColor + ";", "6px 6px"];
        case "dkUpDiag":
            return ["repeating-linear-gradient(-45deg, transparent 1px , #" + bgColor + " 5px)" + "#" + fgColor + ";"];
        case "ltUpDiag":
            return ["repeating-linear-gradient(-45deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
        case "wdDnDiag":
            return ["repeating-linear-gradient(45deg, transparent 1px , transparent 4px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
        case "dkDnDiag":
            return ["repeating-linear-gradient(45deg, transparent 1px , #" + bgColor + " 5px)" + "#" + fgColor + ";"];
        case "ltDnDiag":
            return ["repeating-linear-gradient(45deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
        case "dkHorz":
            return ["repeating-linear-gradient(0deg, transparent 1px , transparent 2px, #" + bgColor + " 7px)" + "#" + fgColor + ";"];
        case "ltHorz":
            return ["repeating-linear-gradient(0deg, transparent 1px , transparent 5px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
        case "narHorz":
            return ["repeating-linear-gradient(0deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
        case "dkVert":
            return ["repeating-linear-gradient(90deg, transparent 1px , transparent 2px, #" + bgColor + " 7px)" + "#" + fgColor + ";"];
        case "ltVert":
            return ["repeating-linear-gradient(90deg, transparent 1px , transparent 5px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
            break
        case "narVert":
            return ["repeating-linear-gradient(90deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
        case "lgCheck":
        case "smCheck": {
            let size = "";
            let pos = "";
            if (prst == "lgCheck") {
                size = "8px 8px";
                pos = "0 0, 4px 4px, 4px 4px, 8px 8px";
            } else {
                size = "4px 4px";
                pos = "0 0, 2px 2px, 2px 2px, 4px 4px";
            }
            return ["linear-gradient(45deg,  #" + fgColor + " 25%, transparent 0, transparent 75%,  #" + fgColor + " 0), " +
                "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 0, transparent 75%,  #" + fgColor + " 0) " +
                "#" + bgColor + ";", size, pos];
        }
        // case "smCheck":
        //     return ["linear-gradient(45deg, transparent 0%, transparent calc(50% - 0.5px),  #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%), " +
        //         "linear-gradient(-45deg, transparent 0%, transparent calc(50% - 0.5px) , #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%)  " +
        //         "#" + bgColor + ";", "4px 4px"];
        //     break 
        case "dashUpDiag":
            return ["repeating-linear-gradient(152deg, #" + fgColor + ", #" + fgColor + " 5% , transparent 0, transparent 70%)" +
                "#" + bgColor + ";", "4px 4px"];
        case "dashDnDiag":
            return ["repeating-linear-gradient(45deg, #" + fgColor + ", #" + fgColor + " 5% , transparent 0, transparent 70%)" +
                "#" + bgColor + ";", "4px 4px"];
        case "diagBrick":
            return ["linear-gradient(45deg, transparent 15%,  #" + fgColor + " 30%, transparent 30%), " +
                "linear-gradient(-45deg, transparent 15%,  #" + fgColor + " 30%, transparent 30%), " +
                "linear-gradient(-45deg, transparent 65%,  #" + fgColor + " 80%, transparent 0) " +
                "#" + bgColor + ";", "4px 4px"];
        case "horzBrick":
            return ["linear-gradient(335deg, #" + bgColor + " 1.6px, transparent 1.6px), " +
                "linear-gradient(155deg, #" + bgColor + " 1.6px, transparent 1.6px), " +
                "linear-gradient(335deg, #" + bgColor + " 1.6px, transparent 1.6px), " +
                "linear-gradient(155deg, #" + bgColor + " 1.6px, transparent 1.6px) " +
                "#" + fgColor + ";", "4px 4px", "0 0.15px, 0.3px 2.5px, 2px 2.15px, 2.35px 0.4px"];

        case "dashVert":
            return ["linear-gradient(0deg,  #" + bgColor + " 30%, transparent 30%)," +
                "linear-gradient(90deg,transparent, transparent 40%, #" + fgColor + " 40%, #" + fgColor + " 60% , transparent 60%)" +
                "#" + bgColor + ";", "4px 4px"];
        case "dashHorz":
            return ["linear-gradient(90deg,  #" + bgColor + " 30%, transparent 30%)," +
                "linear-gradient(0deg,transparent, transparent 40%, #" + fgColor + " 40%, #" + fgColor + " 60% , transparent 60%)" +
                "#" + bgColor + ";", "4px 4px"];
        case "solidDmnd":
            return ["linear-gradient(135deg,  #" + fgColor + " 25%, transparent 25%), " +
                "linear-gradient(225deg,  #" + fgColor + " 25%, transparent 25%), " +
                "linear-gradient(315deg,  #" + fgColor + " 25%, transparent 25%), " +
                "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 25%) " +
                "#" + bgColor + ";", "8px 8px"];
        case "openDmnd":
            return ["linear-gradient(45deg, transparent 0%, transparent calc(50% - 0.5px),  #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%), " +
                "linear-gradient(-45deg, transparent 0%, transparent calc(50% - 0.5px) , #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%) " +
                "#" + bgColor + ";", "8px 8px"];

        case "dotDmnd":
            return ["radial-gradient(#" + fgColor + " 15%, transparent 0), " +
                "radial-gradient(#" + fgColor + " 15%, transparent 0) " +
                "#" + bgColor + ";", "4px 4px", "0 0, 2px 2px"];
        case "zigZag":
        case "wave": {
            let size = "";
            if (prst == "zigZag") size = "0";
            else size = "1px";
            return ["linear-gradient(135deg,  #" + fgColor + " 25%, transparent 25%) 50px " + size + ", " +
                "linear-gradient(225deg,  #" + fgColor + " 25%, transparent 25%) 50px " + size + ", " +
                "linear-gradient(315deg,  #" + fgColor + " 25%, transparent 25%), " +
                "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 25%) " +
                "#" + bgColor + ";", "4px 4px"];
        }
        case "lgConfetti":
        case "smConfetti": {
            let size = "";
            if (prst == "lgConfetti") size = "4px 4px";
            else size = "2px 2px";
            return ["linear-gradient(135deg,  #" + fgColor + " 25%, transparent 25%) 50px 1px, " +
                "linear-gradient(225deg,  #" + fgColor + " 25%, transparent 25%), " +
                "linear-gradient(315deg,  #" + fgColor + " 25%, transparent 25%) 50px 1px , " +
                "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 25%) " +
                "#" + bgColor + ";", size];
        }
        // case "weave":
        //     return ["linear-gradient(45deg,  #" + bgColor + " 5%, transparent 25%) 50px 0, " +
        //         "linear-gradient(135deg,  #" + bgColor + " 25%, transparent 25%) 50px 0, " +
        //         "linear-gradient(45deg,  #" + bgColor + " 25%, transparent 25%) " +
        //         "#" + fgColor + ";", "4px 4px"];
        //     //background: linear-gradient(45deg, #dca 12%, transparent 0, transparent 88%, #dca 0),
        //     //linear-gradient(135deg, transparent 37 %, #a85 0, #a85 63 %, transparent 0),
        //     //linear-gradient(45deg, transparent 37 %, #dca 0, #dca 63 %, transparent 0) #753;
        //     // background-size: 25px 25px;
        //     break;

        case "plaid":
            return ["linear-gradient(0deg, transparent, transparent 25%, #" + fgColor + "33 25%, #" + fgColor + "33 50%)," +
                "linear-gradient(90deg, transparent, transparent 25%, #" + fgColor + "66 25%, #" + fgColor + "66 50%) " +
                "#" + bgColor + ";", "4px 4px"];
            /**
                background-color: #6677dd;
                background-image: 
                repeating-linear-gradient(0deg, transparent, transparent 35px, rgba(255, 255, 255, 0.2) 35px, rgba(255, 255, 255, 0.2) 70px), 
                repeating-linear-gradient(90deg, transparent, transparent 35px, rgba(255,255,255,0.4) 35px, rgba(255,255,255,0.4) 70px);
             */
        case "sphere":
            return ["radial-gradient(#" + fgColor + " 50%, transparent 50%)," +
                "#" + bgColor + ";", "4px 4px"];
        case "weave":
        case "shingle":
            return ["linear-gradient(45deg, #" + bgColor + " 1.31px , #" + fgColor + " 1.4px, #" + fgColor + " 1.5px, transparent 1.5px, transparent 4.2px, #" + fgColor + " 4.2px, #" + fgColor + " 4.3px, transparent 4.31px), " +
                "linear-gradient(-45deg,  #" + bgColor + " 1.31px , #" + fgColor + " 1.4px, #" + fgColor + " 1.5px, transparent 1.5px, transparent 4.2px, #" + fgColor + " 4.2px, #" + fgColor + " 4.3px, transparent 4.31px) 0 4px, " +
                "#" + bgColor + ";", "4px 8px"];
        //background:
        //linear-gradient(45deg, #708090 1.31px, #d9ecff 1.4px, #d9ecff 1.5px, transparent 1.5px, transparent 4.2px, #d9ecff 4.2px, #d9ecff 4.3px, transparent 4.31px),
        //linear-gradient(-45deg, #708090 1.31px, #d9ecff 1.4px, #d9ecff 1.5px, transparent 1.5px, transparent 4.2px, #d9ecff 4.2px, #d9ecff 4.3px, transparent 4.31px)0 4px;
        //background-color:#708090;
        //background-size: 4px 8px;
        case "pct5":
        case "pct10":
        case "pct20":
        case "pct25":
        case "pct30":
        case "pct40":
        case "pct50":
        case "pct60":
        case "pct70":
        case "pct75":
        case "pct80":
        case "pct90":
        //case "dotDmnd":
        case "trellis":
        case "divot":
            let px_pr_ary;
            switch (prst) {
                case "pct5":
                    px_pr_ary = ["0.3px", "10%", "2px 2px"];
                    break
                case "divot":
                    px_pr_ary = ["0.3px", "40%", "4px 4px"];
                    break
                case "pct10":
                    px_pr_ary = ["0.3px", "20%", "2px 2px"];
                    break
                case "pct20":
                    //case "dotDmnd":
                    px_pr_ary = ["0.2px", "40%", "2px 2px"];
                    break
                case "pct25":
                    px_pr_ary = ["0.2px", "50%", "2px 2px"];
                    break
                case "pct30":
                    px_pr_ary = ["0.5px", "50%", "2px 2px"];
                    break
                case "pct40":
                    px_pr_ary = ["0.5px", "70%", "2px 2px"];
                    break
                case "pct50":
                    px_pr_ary = ["0.09px", "90%", "2px 2px"];
                    break
                case "pct60":
                    px_pr_ary = ["0.3px", "90%", "2px 2px"];
                    break
                case "pct70":
                case "trellis":
                    px_pr_ary = ["0.5px", "95%", "2px 2px"];
                    break
                case "pct75":
                    px_pr_ary = ["0.65px", "100%", "2px 2px"];
                    break
                case "pct80":
                    px_pr_ary = ["0.85px", "100%", "2px 2px"];
                    break
                case "pct90":
                    px_pr_ary = ["1px", "100%", "2px 2px"];
                    break
            }
            return ["radial-gradient(#" + fgColor + " " + px_pr_ary[0] + ", transparent " + px_pr_ary[1] + ")," +
                "#" + bgColor + ";", px_pr_ary[2]];
        default:
            return [0, 0];
    }
}

export const shapeGear = (w: number, h: number, points: number) => {
    let innerRadius = h;//gear.innerRadius;
    let outerRadius = 1.5 * innerRadius;
    let cx = outerRadius,//Math.max(innerRadius, outerRadius),                   // center x
        cy = outerRadius,//Math.max(innerRadius, outerRadius),                    // center y
        notches = points,//gear.points,                      // num. of notches
        radiusO = outerRadius,                    // outer radius
        radiusI = innerRadius,                    // inner radius
        taperO = 50,                     // outer taper %
        taperI = 35,                     // inner taper %

        // pre-calculate values for loop

        pi2 = 2 * Math.PI,            // cache 2xPI (360deg)
        angle = pi2 / (notches * 2),    // angle between notches
        taperAI = angle * taperI * 0.005, // inner taper offset (100% = half notch)
        taperAO = angle * taperO * 0.005, // outer taper offset
        a = angle,                  // iterator (angle)
        toggle = false;
    // move to starting point
    let d = " M" + (cx + radiusO * Math.cos(taperAO)) + " " + (cy + radiusO * Math.sin(taperAO));

    // loop
    for (; a <= pi2 + angle; a += angle) {
        // draw inner to outer line
        if (toggle) {
            d += " L" + (cx + radiusI * Math.cos(a - taperAI)) + "," + (cy + radiusI * Math.sin(a - taperAI));
            d += " L" + (cx + radiusO * Math.cos(a + taperAO)) + "," + (cy + radiusO * Math.sin(a + taperAO));
        } else { // draw outer to inner line
            d += " L" + (cx + radiusO * Math.cos(a - taperAO)) + "," + (cy + radiusO * Math.sin(a - taperAO)); // outer line
            d += " L" + (cx + radiusI * Math.cos(a + taperAI)) + "," + (cy + radiusI * Math.sin(a + taperAI));// inner line

        }
        // switch level
        toggle = !toggle;
    }
    // close the final line
    d += " ";
    return d;
}
export const shapeArc = (cX: number, cY: number, rX: number, rY: number, stAng: number, endAng: number, isClose: boolean): string => {
    let dData;
    let angle = stAng;
    if (endAng >= stAng) {
        while (angle <= endAng) {
            let radians = angle * (Math.PI / 180);  // convert degree to radians
            let x = cX + Math.cos(radians) * rX;
            let y = cY + Math.sin(radians) * rY;
            if (angle == stAng) {
                dData = " M" + x + " " + y;
            }
            dData += " L" + x + " " + y;
            angle++;
        }
    } else {
        while (angle > endAng) {
            let radians = angle * (Math.PI / 180);  // convert degree to radians
            let x = cX + Math.cos(radians) * rX;
            let y = cY + Math.sin(radians) * rY;
            if (angle == stAng) {
                dData = " M " + x + " " + y;
            }
            dData += " L " + x + " " + y;
            angle--;
        }
    }
    dData += (isClose ? " z" : "");
    return dData!;
}

export const shapeSnipRoundRect = (w: number, h: number, adj1: number, adj2: number, shapeType: string, adjType: string | undefined) => {
    /* 
    shapeType: snip,round
    adjType: cornr1,cornr2,cornrAll,diag
    */
    let adjA = 0, adjB = 0, adjC = 0, adjD = adj1;
    if (adjType == "cornr1") {
        adjA = 0;
        adjB = 0;
        adjC = 0;
        adjD = adj1;
    } else if (adjType == "cornr2") {
        adjA = adj1;
        adjB = adj2;
        adjC = adj2;
        adjD = adj1;
    } else if (adjType == "cornrAll") {
        adjA = adj1;
        adjB = adj1;
        adjC = adj1;
        adjD = adj1;
    } else if (adjType == "diag") {
        adjA = adj1;
        adjB = adj2;
        adjC = adj1;
        adjD = adj2;
    }
    //d is a string that describes the path of the slice.
    let d;
    if (shapeType == "round") {
        d = "M0" + "," + (h / 2 + (1 - adjB) * (h / 2)) + " Q" + 0 + "," + h + " " + adjB * (w / 2) + "," + h + " L" + (w / 2 + (1 - adjC) * (w / 2)) + "," + h +
            " Q" + w + "," + h + " " + w + "," + (h / 2 + (h / 2) * (1 - adjC)) + "L" + w + "," + (h / 2) * adjD +
            " Q" + w + "," + 0 + " " + (w / 2 + (w / 2) * (1 - adjD)) + ",0 L" + (w / 2) * adjA + ",0" +
            " Q" + 0 + "," + 0 + " 0," + (h / 2) * (adjA) + " z";
    } else if (shapeType == "snip") {
        d = "M0" + "," + adjA * (h / 2) + " L0" + "," + (h / 2 + (h / 2) * (1 - adjB)) + "L" + adjB * (w / 2) + "," + h +
            " L" + (w / 2 + (w / 2) * (1 - adjC)) + "," + h + "L" + w + "," + (h / 2 + (h / 2) * (1 - adjC)) +
            " L" + w + "," + adjD * (h / 2) + "L" + (w / 2 + (w / 2) * (1 - adjD)) + ",0 L" + ((w / 2) * adjA) + ",0 z";
    }
    return d;
}

export const shapePie = (H: number | undefined, w: number, adj1: number | undefined, adj2: number | undefined, isClose: boolean | undefined) =>  {
    // @ts-ignore
    let pieVal = parseInt(adj2);
    // @ts-ignore
    let piAngle = parseInt(adj1);
    // @ts-ignore
    let size = parseInt(H),
        radius = (size / 2),
        value = pieVal - piAngle;
    if (value < 0) {
        value = 360 + value;
    }
    //console.log("value: ",value)      
    value = Math.min(Math.max(value, 0), 360);

    //calculate x,y coordinates of the point on the circle to draw the arc to. 
    let x = Math.cos((2 * Math.PI) / (360 / value));
    let y = Math.sin((2 * Math.PI) / (360 / value));


    //d is a string that describes the path of the slice.
    let longArc, d, rot;
    if (isClose) {
        longArc = (value <= 180) ? 0 : 1;
        d = "M" + radius + "," + radius + " L" + radius + "," + 0 + " A" + radius + "," + radius + " 0 " + longArc + ",1 " + (radius + y * radius) + "," + (radius - x * radius) + " z";
        rot = "rotate(" + (piAngle - 270) + ", " + radius + ", " + radius + ")";
    } else {
        longArc = (value <= 180) ? 0 : 1;
        let radius1 = radius;
        let radius2 = w / 2;
        d = "M" + radius1 + "," + 0 + " A" + radius2 + "," + radius1 + " 0 " + longArc + ",1 " + (radius2 + y * radius2) + "," + (radius1 - x * radius1);
        rot = "rotate(" + (piAngle + 90) + ", " + radius + ", " + radius + ")";
    }

    return [d, rot];
}

export const getSvgGradient = (w: number, h: number, angl: number, color_arry: string[], shpId: string) => {
    // TODO: find meaning
    // @ts-ignore
    let stopsArray = getMiddleStops(color_arry - 2);

    let svgAngle = '',
        svgHeight = h,
        svgWidth = w,
        svg = '',
        xy_ary = SVGangle(angl, svgHeight, svgWidth),
        x1 = xy_ary[0],
        y1 = xy_ary[1],
        x2 = xy_ary[2],
        y2 = xy_ary[3];

    let sal = stopsArray.length,
        sr = sal < 20 ? 100 : 1000;
    svgAngle = ' gradientUnits="userSpaceOnUse" x1="' + x1 + '%" y1="' + y1 + '%" x2="' + x2 + '%" y2="' + y2 + '%"';
    svgAngle = '<linearGradient id="linGrd_' + shpId + '"' + svgAngle + '>\n';
    svg += svgAngle;

    for (var i = 0; i < sal; i++) {
        let tinClr = tinycolor("#" + color_arry[i]);
        let alpha = tinClr.getAlpha();
        //console.log("color: ", color_arry[i], ", rgba: ", tinClr.toHexString(), ", alpha: ", alpha)
        svg += '<stop offset="' + Math.round(parseFloat(stopsArray[i]) / 100 * sr) / sr + '" style="stop-color:' + tinClr.toHexString() + '; stop-opacity:' + (alpha) + ';"';
        svg += '/>\n'
    }

    svg += '</linearGradient>\n' + '';

    return svg
}
