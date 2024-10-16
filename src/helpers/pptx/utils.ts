const archaicNumbers = (arr: [number | RegExp, string][]) => {
    const arrParse = arr.slice().sort((a, b) => b[1].length - a[1].length);
    return {
        format: (n: number): string => {
            let ret = '';
            for (const [source, replacement] of arrParse) {
                if (typeof source === 'number') {
                    while (n >= source) {
                        ret += replacement;
                        n -= source;
                    }
                } else {
                    ret = ret.replace(source, replacement);
                }
            }
            return ret;
        }
    };
}

const hebrew2Minus = archaicNumbers([
    [1000, ''],
    [400, 'ת'],
    [300, 'ש'],
    [200, 'ר'],
    [100, 'ק'],
    [90, 'צ'],
    [80, 'פ'],
    [70, 'ע'],
    [60, 'ס'],
    [50, 'נ'],
    [40, 'מ'],
    [30, 'ל'],
    [20, 'כ'],
    [10, 'י'],
    [9, 'ט'],
    [8, 'ח'],
    [7, 'ז'],
    [6, 'ו'],
    [5, 'ה'],
    [4, 'ד'],
    [3, 'ג'],
    [2, 'ב'],
    [1, 'א'],
    [/יה/, 'ט״ו'],
    [/יו/, 'ט״ז'],
    [/([א-ת])([א-ת])$/, '$1״$2'],
    [/^([א-ת])$/, "$1׳"]
]);

export const getNumTypeNum = (numTyp: string, num: number) => {
    let rtrnNum = "";
    switch (numTyp) {
        case "arabicPeriod":
            rtrnNum = num + ". ";
            break;
        case "arabicParenR":
            rtrnNum = num + ") ";
            break;
        case "alphaLcParenR":
            rtrnNum = alphaNumeric(num, "lowerCase") + ") ";
            break;
        case "alphaLcPeriod":
            rtrnNum = alphaNumeric(num, "lowerCase") + ". ";
            break;

        case "alphaUcParenR":
            rtrnNum = alphaNumeric(num, "upperCase") + ") ";
            break;
        case "alphaUcPeriod":
            rtrnNum = alphaNumeric(num, "upperCase") + ". ";
            break;

        case "romanUcPeriod":
            rtrnNum = romanize(num) + ". ";
            break;
        case "romanLcParenR":
            rtrnNum = romanize(num) + ") ";
            break;
        case "hebrew2Minus":
            rtrnNum = hebrew2Minus.format(num) + "-";
            break;
        default:
            rtrnNum = `${num}`;
    }
    return rtrnNum;
}

const romanizeValues = ["", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM",
    "", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC",
    "", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX"];

export const romanize = (num: number) => {
    if (!+num)
        return false;
    let digits = String(+num).split(""),
        roman = "",
        i = 3;
    while (i--)
        roman = (romanizeValues[+digits.pop()! + (i * 10)] || "") + roman;
    return Array(+digits.join("") + 1).join("M") + roman;
}

export const alphaNumeric = (num: number, upperLower: string) => {
    num = Number(num) - 1;
    let aNum = "";
    if (upperLower == "upperCase") {
        aNum = (((num / 26 >= 1) ? String.fromCharCode(num / 26 + 64) : '') + String.fromCharCode(num % 26 + 65)).toUpperCase();
    } else if (upperLower == "lowerCase") {
        aNum = (((num / 26 >= 1) ? String.fromCharCode(num / 26 + 64) : '') + String.fromCharCode(num % 26 + 65)).toLowerCase();
    }
    return aNum;
}

const base64Encodings = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

export const base64ArrayBuffer = (arrayBuffer: ArrayBuffer) => {
    let base64 = '';
    let bytes = new Uint8Array(arrayBuffer);
    let byteLength = bytes.byteLength;
    let byteRemainder = byteLength % 3;
    let mainLength = byteLength - byteRemainder;

    let a, b, c, d;
    let chunk;

    for (var i = 0; i < mainLength; i = i + 3) {
        chunk = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
        a = (chunk & 16515072) >> 18;
        b = (chunk & 258048) >> 12;
        c = (chunk & 4032) >> 6;
        d = chunk & 63;
        base64 += base64Encodings[a] + base64Encodings[b] + base64Encodings[c] + base64Encodings[d];
    }

    if (byteRemainder == 1) {
        chunk = bytes[mainLength];
        a = (chunk & 252) >> 2;
        b = (chunk & 3) << 4;
        base64 += base64Encodings[a] + base64Encodings[b] + '==';
    } else if (byteRemainder == 2) {
        chunk = (bytes[mainLength] << 8) | bytes[mainLength + 1];
        a = (chunk & 64512) >> 10;
        b = (chunk & 1008) >> 4;
        c = (chunk & 15) << 2;
        base64 += base64Encodings[a] + base64Encodings[b] + base64Encodings[c] + '=';
    }

    return base64;
}

const htmlStringMap: {[key: string]: string} = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;',
    '\t': '&nbsp;&nbsp;&nbsp;&nbsp;',
    '\s': '&nbsp;'
};

export const escapeHtml = (text: string) => {
    return text.replace(/[&<>"'\t\s]|/g, (match) => htmlStringMap[match] || match);
}
