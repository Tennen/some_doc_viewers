const urlRegex = /^(https?|ftp):\/\/([a-zA-Z0-9.-]+(:[a-zA-Z0-9.&%$-]+)*@)*((25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9][0-9]?)(\.(25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9]?[0-9])){3}|([a-zA-Z0-9-]+\.)*[a-zA-Z0-9-]+\.(com|edu|gov|int|mil|net|org|biz|arpa|info|name|pro|aero|coop|museum|[a-zA-Z]{2}))(:[0-9]+)*(\/($|[a-zA-Z0-9.,?'\\+&%$#=~_-]+))*$/;

export const isVideoLink = (vdoFile: string) => {
    return urlRegex.test(vdoFile);
}

export const extractFileExtension = (filename: string) => {
    return filename?.slice(filename?.lastIndexOf(".") + 1)?.toLowerCase();
}

export const getMimeType = (imgFileExt: string | undefined) => {
    let mimeType = "";
    //console.log(imgFileExt)
    switch (imgFileExt?.toLowerCase()) {
        case "jpg":
        case "jpeg":
            mimeType = "image/jpeg";
            break;
        case "png":
            mimeType = "image/png";
            break;
        case "gif":
            mimeType = "image/gif";
            break;
        case "emf": // Not native support
            mimeType = "image/x-emf";
            break;
        case "wmf": // Not native support
            mimeType = "image/x-wmf";
            break;
        case "svg":
            mimeType = "image/svg+xml";
            break;
        case "mp4":
            mimeType = "video/mp4";
            break;
        case "webm":
            mimeType = "video/webm";
            break;
        case "ogg":
            mimeType = "video/ogg";
            break;
        case "avi":
            mimeType = "video/avi";
            break;
        case "mpg":
            mimeType = "video/mpg";
            break;
        case "wmv":
            mimeType = "video/wmv";
            break;
        case "mp3":
            mimeType = "audio/mpeg";
            break;
        case "wav":
            mimeType = "audio/wav";
            break;
        // case "emf":
        //     mimeType = "image/emf";
        //     break;
        // case "wmf":
        //     mimeType = "image/wmf";
        //     break;
        case "tif":
        case "tiff":
            mimeType = "image/tiff";
            break;
    }
    return mimeType;
}
