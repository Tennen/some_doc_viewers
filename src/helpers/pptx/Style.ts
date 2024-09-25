interface StyleObject { name: string; text: string }

export class StyleTable {
    map: Record<string, StyleObject> = {};

    get(key: string) {
        return this.map[key];
    }

    update(key: string, item: StyleObject) {
        this.map[key] = item;
    }

    nextIndex() {
        return Object.keys(this.map).length + 1;
    }

    has(key: string) {
        return key in this.map;
    }

    genGlobalCSS() {
        let cssText = "";
        for (var key in this.map) {
            let tagname = "";
            //TODO: ADD suffix
            cssText += tagname + " ." + this.map[key]["name"] + "{" + this.map[key]["text"] + "}\n";
        }
        //cssText += " .slide{margin-bottom: 5px;}\n"; // TODO

        return cssText;
    }
}