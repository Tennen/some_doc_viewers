import { dingbatUnicode } from './dingbatUnicode';

const htmlStringMap: {[key: string]: string} = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;',
    '\t': '&nbsp;&nbsp;&nbsp;&nbsp;',
    '\s': '&nbsp;',
    ' ': '&nbsp;'
};

export const escapeHtml = (text: string) => {
    return text.replace(/[&<>"'\t\s]|/g, (match) => htmlStringMap[match] || match);
}

export const estimateWidth = (prgr_text: string) => {
    const template = document.createElement('div');
    template.innerHTML = prgr_text;
    template.style.position = 'absolute';
    template.style.whiteSpace = 'nowrap';
    template.style.visibility = 'hidden';
    document.body.appendChild(template);
    const width = template.offsetWidth;
    document.body.removeChild(template);
    return width;
}

export const getHtmlBullet = (typefaceNode: any, buChar: string) => {
    //http://www.alanwood.net/demos/wingdings.html
    //not work for IE11
    //console.log("genBuChar typefaceNode:", typefaceNode, " buChar:", buChar, "charCodeAt:", buChar.charCodeAt(0))
    switch (buChar) {
        case "§":
            return "&#9632;";//"■"; //9632 | U+25A0 | Black square
            break;
        case "q":
            return "&#10065;";//"❑"; // 10065 | U+2751 | Lower right shadowed white square
            break;
        case "v":
            return "&#10070;";//"❖"; //10070 | U+2756 | Black diamond minus white X
            break;
        case "Ø":
            return "&#11162;";//"⮚"; //11162 | U+2B9A | Three-D top-lighted rightwards equilateral arrowhead
            break;
        case "ü":
            return "&#10004;";//"✔";  //10004 | U+2714 | Heavy check mark
            break;
        default:
            if (/*typefaceNode == "Wingdings" ||*/ typefaceNode == "Wingdings 2" || typefaceNode == "Wingdings 3") {
                let wingCharCode = getDingbatToUnicode(typefaceNode, buChar);
                if (wingCharCode !== null) {
                    return "&#" + wingCharCode + ";";
                }
            }
            return "&#" + (buChar.charCodeAt(0)) + ";";
    }
}
const getDingbatToUnicode = (typefaceNode: any, buChar: string) => {
    let dingbatCode = buChar.codePointAt(0)! & 0xFFF;
    let char_unicode = null;
    let len = dingbatUnicode.length;
    let i = 0;
    while (len--) {
        // blah blah
        let item = dingbatUnicode[i];
        if (item.f == typefaceNode && Number(item.code) == dingbatCode) {
            char_unicode = item.unicode;
            break;
        }
        i++;
    }
    return char_unicode
}
