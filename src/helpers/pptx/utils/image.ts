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

export const getBase64ImageDimensions = (imgSrc: string) => {
    let image = new Image();
    let w, h;
    image.onload = function () {
        w = image.width;
        h = image.height;
    };
    image.src = imgSrc;

    do {
        if (image.width !== undefined) {
            return [image.width, image.height];
        }
    } while (image.width === undefined);

    //return [w, h];
}
