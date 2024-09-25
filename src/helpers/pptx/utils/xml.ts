import { XMLParser } from 'fast-xml-parser';

const xmlParser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "",
    attributesGroupName: "attrs",
    ignorePiTags: true,
    trimValues: false,
    cdataPropName: '@CDATA',
});

export const parse = (xml?: string) => {
    if (!xml) {
        return null;
    }

    const result = xmlParser.parse(xml);

    return result;
};
