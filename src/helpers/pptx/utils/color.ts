import tinycolor from 'tinycolor2';
import * as _ from 'lodash';

// ===== Color functions =====
/**
 * applyShade
 * @param {string} rgbStr
 * @param {number} shadeValue
 */
export const applyShade = (rgbStr: string, shadeValue: number, isAlpha: boolean): string => {
    const color = tinycolor(rgbStr).toHsl();
    const adjustedShade = Math.min(shadeValue, 1);
    const calcL = Math.min(color.l * adjustedShade, 1);
    
    return isAlpha
        ? tinycolor({ h: color.h, s: color.s, l: calcL, a: color.a }).toHex8()
        : tinycolor({ h: color.h, s: color.s, l: calcL, a: color.a }).toHex();
};

/**
 * applyTint
 * @param {string} rgbStr
 * @param {number} tintValue
 */
export const applyTint = (rgbStr: string, tintValue: number, isAlpha: boolean): string => {
    const color = tinycolor(rgbStr).toHsl();
    const adjustedTint = Math.min(tintValue, 1);
    const calcL = color.l * adjustedTint + (1 - adjustedTint);
    
    return isAlpha
        ? tinycolor({ h: color.h, s: color.s, l: calcL, a: color.a }).toHex8()
        : tinycolor({ h: color.h, s: color.s, l: calcL, a: color.a }).toHex();
};

/**
 * applyLumOff
 * @param {string} rgbStr
 * @param {number} offset
 */
export const applyLumOff = (rgbStr: string, offset: number, isAlpha: boolean): string => {
    const color = tinycolor(rgbStr).toHsl();
    const lum = Math.min(offset + color.l, 1);
    
    return isAlpha
        ? tinycolor({ h: color.h, s: color.s, l: lum, a: color.a }).toHex8()
        : tinycolor({ h: color.h, s: color.s, l: lum, a: color.a }).toHex();
};

/**
 * applyLumMod
 * @param {string} rgbStr
 * @param {number} multiplier
 */
export const applyLumMod = (rgbStr: string, multiplier: number, isAlpha: boolean): string => {
    const color = tinycolor(rgbStr).toHsl();
    const calcL = Math.min(color.l * multiplier, 1);
    
    return isAlpha
        ? tinycolor({ h: color.h, s: color.s, l: calcL, a: color.a }).toHex8()
        : tinycolor({ h: color.h, s: color.s, l: calcL, a: color.a }).toHex();
};

/**
 * applyHueMod
 * @param {string} rgbStr
 * @param {number} multiplier
 */
export const applyHueMod = (rgbStr: string, multiplier: number, isAlpha: boolean): string => {
    const color = tinycolor(rgbStr).toHsl();
    let calcH = color.h * multiplier;
    calcH = calcH >= 360 ? calcH - 360 : calcH;
    
    return isAlpha
        ? tinycolor({ h: calcH, s: color.s, l: color.l, a: color.a }).toHex8()
        : tinycolor({ h: calcH, s: color.s, l: color.l, a: color.a }).toHex();
};

/**
 * applySatMod
 * @param {string} rgbStr
 * @param {number} multiplier
 */
export const applySatMod = (rgbStr: string, multiplier: number, isAlpha: boolean): string => {
    const color = tinycolor(rgbStr).toHsl();
    const calcS = Math.min(color.s * multiplier, 1);
    
    return isAlpha
        ? tinycolor({ h: color.h, s: calcS, l: color.l, a: color.a }).toHex8()
        : tinycolor({ h: color.h, s: calcS, l: color.l, a: color.a }).toHex();
};

/**
 * rgba2hex
 * @param {string} rgbaStr
 */
export const rgba2hex = (rgbaStr: string): string => {
    const rgb = rgbaStr.replace(/\s/g, '').match(/^rgba?\((\d+),(\d+),(\d+),?([^,\s)]+)?/i);
    const alpha = (rgb?.[4] || "").trim();
    
    if (!rgb) return rgbaStr;
    
    const hex = (Number(rgb[1]) | 1 << 8).toString(16).slice(1) +
                (Number(rgb[2]) | 1 << 8).toString(16).slice(1) +
                (Number(rgb[3]) | 1 << 8).toString(16).slice(1);
    
    const a = alpha !== "" ? alpha : "1";
    const hexAlpha = ((Number(a) * 255) | 1 << 8).toString(16).slice(1);
    
    return hex + hexAlpha;
};

export const angleToDegrees = (angle: string | number | null): number => {
    if (_.isEmpty(angle) || _.isNil(angle)) {
        return 0;
    }
    return Math.round(Number(angle) / 60000);
};

export const colorMap = new Map<string, string>([
    ['white', 'ffffff'], ['AliceBlue', 'f0f8ff'], ['AntiqueWhite', 'faebd7'],
    ['Aqua', '00ffff'], ['Aquamarine', '7fffd4'], ['Azure', 'f0ffff'],
    ['Beige', 'f5f5dc'], ['Bisque', 'ffe4c4'], ['black', '000000'],
    ['BlanchedAlmond', 'ffebcd'], ['Blue', '0000ff'], ['BlueViolet', '8a2be2'],
    ['Brown', 'a52a2a'], ['BurlyWood', 'deb887'], ['CadetBlue', '5f9ea0'],
    ['Chartreuse', '7fff00'], ['Chocolate', 'd2691e'], ['Coral', 'ff7f50'],
    ['CornflowerBlue', '6495ed'], ['Cornsilk', 'fff8dc'], ['Crimson', 'dc143c'],
    ['Cyan', '00ffff'], ['DarkBlue', '00008b'], ['DarkCyan', '008b8b'],
    ['DarkGoldenRod', 'b8860b'], ['DarkGray', 'a9a9a9'], ['DarkGrey', 'a9a9a9'],
    ['DarkGreen', '006400'], ['DarkKhaki', 'bdb76b'], ['DarkMagenta', '8b008b'],
    ['DarkOliveGreen', '556b2f'], ['DarkOrange', 'ff8c00'], ['DarkOrchid', '9932cc'],
    ['DarkRed', '8b0000'], ['DarkSalmon', 'e9967a'], ['DarkSeaGreen', '8fbc8f'],
    ['DarkSlateBlue', '483d8b'], ['DarkSlateGray', '2f4f4f'], ['DarkSlateGrey', '2f4f4f'],
    ['DarkTurquoise', '00ced1'], ['DarkViolet', '9400d3'], ['DeepPink', 'ff1493'],
    ['DeepSkyBlue', '00bfff'], ['DimGray', '696969'], ['DimGrey', '696969'],
    ['DodgerBlue', '1e90ff'], ['FireBrick', 'b22222'], ['FloralWhite', 'fffaf0'],
    ['ForestGreen', '228b22'], ['Fuchsia', 'ff00ff'], ['Gainsboro', 'dcdcdc'],
    ['GhostWhite', 'f8f8ff'], ['Gold', 'ffd700'], ['GoldenRod', 'daa520'],
    ['Gray', '808080'], ['Grey', '808080'], ['Green', '008000'],
    ['GreenYellow', 'adff2f'], ['HoneyDew', 'f0fff0'], ['HotPink', 'ff69b4'],
    ['IndianRed', 'cd5c5c'], ['Indigo', '4b0082'], ['Ivory', 'fffff0'],
    ['Khaki', 'f0e68c'], ['Lavender', 'e6e6fa'], ['LavenderBlush', 'fff0f5'],
    ['LawnGreen', '7cfc00'], ['LemonChiffon', 'fffacd'], ['LightBlue', 'add8e6'],
    ['LightCoral', 'f08080'], ['LightCyan', 'e0ffff'], ['LightGoldenRodYellow', 'fafad2'],
    ['LightGray', 'd3d3d3'], ['LightGrey', 'd3d3d3'], ['LightGreen', '90ee90'],
    ['LightPink', 'ffb6c1'], ['LightSalmon', 'ffa07a'], ['LightSeaGreen', '20b2aa'],
    ['LightSkyBlue', '87cefa'], ['LightSlateGray', '778899'], ['LightSlateGrey', '778899'],
    ['LightSteelBlue', 'b0c4de'], ['LightYellow', 'ffffe0'], ['Lime', '00ff00'],
    ['LimeGreen', '32cd32'], ['Linen', 'faf0e6'], ['Magenta', 'ff00ff'],
    ['Maroon', '800000'], ['MediumAquaMarine', '66cdaa'], ['MediumBlue', '0000cd'],
    ['MediumOrchid', 'ba55d3'], ['MediumPurple', '9370db'], ['MediumSeaGreen', '3cb371'],
    ['MediumSlateBlue', '7b68ee'], ['MediumSpringGreen', '00fa9a'], ['MediumTurquoise', '48d1cc'],
    ['MediumVioletRed', 'c71585'], ['MidnightBlue', '191970'], ['MintCream', 'f5fffa'],
    ['MistyRose', 'ffe4e1'], ['Moccasin', 'ffe4b5'], ['NavajoWhite', 'ffdead'],
    ['Navy', '000080'], ['OldLace', 'fdf5e6'], ['Olive', '808000'],
    ['OliveDrab', '6b8e23'], ['Orange', 'ffa500'], ['OrangeRed', 'ff4500'],
    ['Orchid', 'da70d6'], ['PaleGoldenRod', 'eee8aa'], ['PaleGreen', '98fb98'],
    ['PaleTurquoise', 'afeeee'], ['PaleVioletRed', 'db7093'], ['PapayaWhip', 'ffefd5'],
    ['PeachPuff', 'ffdab9'], ['Peru', 'cd853f'], ['Pink', 'ffc0cb'],
    ['Plum', 'dda0dd'], ['PowderBlue', 'b0e0e6'], ['Purple', '800080'],
    ['RebeccaPurple', '663399'], ['Red', 'ff0000'], ['RosyBrown', 'bc8f8f'],
    ['RoyalBlue', '4169e1'], ['SaddleBrown', '8b4513'], ['Salmon', 'fa8072'],
    ['SandyBrown', 'f4a460'], ['SeaGreen', '2e8b57'], ['SeaShell', 'fff5ee'],
    ['Sienna', 'a0522d'], ['Silver', 'c0c0c0'], ['SkyBlue', '87ceeb'],
    ['SlateBlue', '6a5acd'], ['SlateGray', '708090'], ['SlateGrey', '708090'],
    ['Snow', 'fffafa'], ['SpringGreen', '00ff7f'], ['SteelBlue', '4682b4'],
    ['Tan', 'd2b48c'], ['Teal', '008080'], ['Thistle', 'd8bfd8'],
    ['Tomato', 'ff6347'], ['Turquoise', '40e0d0'], ['Violet', 'ee82ee'],
    ['Wheat', 'f5deb3'], ['White', 'ffffff'], ['WhiteSmoke', 'f5f5f5'],
    ['Yellow', 'ffff00'], ['YellowGreen', '9acd32']
]);

export const rtlLangs = ["he-IL", "ar-AE", "ar-SA", "dv-MV", "fa-IR","ur-PK"];


export const toHex = (n: number) => {
    let hex = n.toString(16);
    while (hex.length < 2) { hex = "0" + hex; }
    return hex;
}

export const hslToRgb = (hue: number, sat: number, light: number) => {
    let t1, t2, r, g, b;
    hue = hue / 60;
    if (light <= 0.5) {
        t2 = light * (sat + 1);
    } else {
        t2 = light + sat - (light * sat);
    }
    t1 = light * 2 - t2;
    r = hueToRgb(t1, t2, hue + 2) * 255;
    g = hueToRgb(t1, t2, hue) * 255;
    b = hueToRgb(t1, t2, hue - 2) * 255;
    return { r: r, g: g, b: b };
}

export const hueToRgb = (t1: number, t2: number, hue: number) => {
    if (hue < 0) hue += 6;
    if (hue >= 6) hue -= 6;
    if (hue < 1) return (t2 - t1) * hue + t1;
    else if (hue < 3) return t2;
    else if (hue < 4) return (t2 - t1) * (4 - hue) + t1;
    else return t1;
}
