import { PPTX } from './PPTX';

export const convertPPTXDataToJSX = (result: any, pptx: PPTX) => {
    if (!result) return;
    let htmlString = "";
    const slideHeight = pptx.basicInfo?.height ?? 0;
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
    console.log(htmlString);
    // if (!options.slideMode || (options.slideMode && options.slideType == "revealjs")) {
        // TODO: implement slideMode
        // if (document.getElementById("all_slides_warpper") === null) {
        //     $("#" + divId + " .slide").wrapAll("<div id='all_slides_warpper' class='slides'></div>");
        //     //$("#" + divId + " .slides").wrap("<div class='reveal'></div>");
        // }

        // if (settings.slideMode && settings.slideType == "revealjs") {
        //     $("#" + divId).addClass("reveal")
        // }
    // }

    //console.log("slidesHeight: " + slidesHeight + "\nnumOfSlides: " + numOfSlides + "\nScale: " + sScaleVal)
    return <div style={{ height: slideHeight * (pptx.basicInfo?.slides.length ?? 0) }} dangerouslySetInnerHTML={{ __html: htmlString }}></div>
};
