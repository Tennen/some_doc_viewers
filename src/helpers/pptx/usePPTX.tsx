import { RefObject, useEffect, useMemo, useRef, useState } from 'react';
import { PPTX } from './PPTX';
import '../../styles/pptx.less';

interface Options {
    url?: string;
    file?: ArrayBuffer;
    containerRef?: RefObject<HTMLElement>;
}

export const usePPTX = (options: Options) => {
    const pptxRef = useRef<PPTX | null>(null);
    useEffect(() => {
        pptxRef.current = new PPTX({
            file: options.file,
            url: options.url,
        });
    }, [])
    useEffect(() => {
        if (!options.containerRef?.current) {
            return;
        }
        pptxRef.current?.render(options.containerRef?.current);
    }, [options.containerRef?.current]);
}
