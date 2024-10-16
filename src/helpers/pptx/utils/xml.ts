import { XMLParser } from 'fast-xml-parser';

export const xmlParser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "",
    attributesGroupName: "attrs",
    ignorePiTags: true,
    trimValues: false,
    cdataPropName: '@CDATA',
});