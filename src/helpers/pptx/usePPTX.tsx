import { RefObject, useEffect, useRef } from 'react';
import { PPTX, PPTXOptions } from './PPTX';
import '../../styles/pptx.less';

interface Options extends PPTXOptions {
    containerRef?: RefObject<HTMLElement>;
}

export const usePPTX = (options: Options) => {
    const pptxRef = useRef<PPTX | null>(null);
    const { containerRef, ...restOptions } = options;
    const loadingRef = useRef(false);
    useEffect(() => {
        pptxRef.current = new PPTX(restOptions);
    }, [])
    useEffect(() => {
        if (!options.containerRef?.current || loadingRef.current) {
            return;
        }
        pptxRef.current?.render(options.containerRef?.current).then(() => {
            loadingRef.current = true;
        });
    }, [options.containerRef?.current]);
}
