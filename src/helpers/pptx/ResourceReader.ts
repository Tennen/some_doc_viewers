import JSZip, { OutputType } from "jszip";
import { parse } from "./utils/xml";
import { base64ArrayBuffer } from "./utils/image";
import { getMimeType } from "./utils/file";

export class ResourceReader {
    private zip: JSZip | null = null;
    private loadingPromise: Promise<void> | null = null;
    constructor(file: ArrayBuffer) {
        this.loadingPromise = this.load(file);
    }

    private async load(file: ArrayBuffer) {
        this.zip = await JSZip.loadAsync(file);
        this.loadingPromise = null;
    }

    async getFile(filename: string, extension: OutputType) {
        if (this.loadingPromise) {
            await this.loadingPromise;
        }
        return await this.zip?.file(filename)?.async(extension);
    }

    async readXmlFile(filename: string) {
        try {
            const fileContent = await this.getFile(filename, "text");
            if (fileContent) {
                return parse(fileContent);
            }
        } catch (e) {
            console.log("error readXmlFile: the file '", filename, "' not exit")
        }
        return null;
    }

    async getImgBase64Data(filename: string) {
        const imgExt = filename.split(".").pop();
        if (imgExt == "xml") {
            return undefined;
        }
        const fileBase64 = await this.getFile(filename, "base64");
        const imgMimeType = getMimeType(imgExt);
        if (fileBase64) return "data:" + imgMimeType + ";base64," + fileBase64;
        return undefined;
    }

}
