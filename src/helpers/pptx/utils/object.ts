import _ from 'lodash';

/**
     * getTextByPathList
     * @param {Object} node
     * @param {string Array} path
     */
export const getTextByPathList = (node: any, path: (string | number)[]) => {

    if (path.constructor !== Array) {
        throw Error("Error of path type! path is not array.");
    }

    if (node === undefined) {
        return undefined;
    }

    let l = path.length;
    for (var i = 0; i < l; i++) {
        node = node[path[i]];
        if (node === undefined) {
            return undefined;
        }
    }

    return node;
}
/**
 * setTextByPathList
 * @param {Object} node
 * @param {string Array} path
 * @param {string} value
 */
export const setTextByPathList = (node: any, path: string[], value: string) => {

    if (path.constructor !== Array) {
        throw Error("Error of path type! path is not array.");
    }

    if (node === undefined) {
        return undefined;
    }

    _.set(node, path, value);
}

/**
 * eachElement
 * @param {Object} node
 * @param {function} doFunction
 */
export const eachElement = (node: any, doFunction: Function) => {
    if (node === undefined) {
        return;
    }
    let result = "";
    if (node.constructor === Array) {
        let l = node.length;
        for (var i = 0; i < l; i++) {
            result += doFunction(node[i], i);
        }
    } else {
        result += doFunction(node, 0);
    }
    return result;
}
