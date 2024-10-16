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
