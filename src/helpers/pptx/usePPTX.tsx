import { useEffect, useMemo, useRef, useState } from 'react';
import { PPTX } from './PPTX';
import { convertPPTXDataToJSX } from './jsxConverter';
import '../../styles/pptx.less';

interface Options {
    url?: string;
    file?: ArrayBuffer;
    slideMode?: boolean;
    slideType?: string;
}

export const usePPTX = (options: Options) => {
    const [pptxRes, setPPTXRes] = useState<any>();
    const pptxInstance = useRef(new PPTX(options));
    useEffect(() => {
        pptxInstance.current.convert().then(result => {
            setPPTXRes(result)
        });
    }, []);

    const JSXResult = useMemo(() => {
        return convertPPTXDataToJSX(pptxRes, pptxInstance.current)
    }, [pptxRes]);

    return JSXResult;
}