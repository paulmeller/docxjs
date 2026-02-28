export function escapeClassName(className: string) {
	return className?.replace(/[ .]+/g, '-').replace(/[&]+/g, 'and').toLowerCase();
}

export function encloseFontFamily(fontFamily: string): string {
    return /^[^"'].*\s.*[^"']$/.test(fontFamily) ? `'${fontFamily}'` : fontFamily;
}

export function splitPath(path: string): [string, string] {
    let si = path.lastIndexOf('/') + 1;
    let folder = si == 0 ? "" : path.substring(0, si);
    let fileName = si == 0 ? path : path.substring(si);

    return [folder, fileName];
}

export function resolvePath(path: string, base: string): string {
    try {
        const prefix = "http://docx/";
        const url = new URL(path, prefix + base).toString();
        return url.substring(prefix.length);
    } catch {
        return `${base}${path}`;
    }
}

export function keyBy<T = any>(array: T[], by: (x: T) => any): Record<any, T> {
    return array.reduce((a, x) => {
        a[by(x)] = x;
        return a;
    }, {});
}

export function blobToBase64(blob: Blob): Promise<string> {
	return new Promise((resolve, reject) => {
		const reader = new FileReader();
		reader.onloadend = () => resolve(reader.result as string);
		reader.onerror = () => reject();
		reader.readAsDataURL(blob);
	});
}

export function isObject(item) {
    return item && typeof item === 'object' && !Array.isArray(item);
}

export function isString(item: unknown): item is string {
    return typeof item === 'string' || item instanceof String;
}

export function mergeDeep(target, ...sources) {
    if (!sources.length) 
        return target;
    
    const source = sources.shift();

    if (isObject(target) && isObject(source)) {
        for (const key in source) {
            if (isObject(source[key])) {
                const val = target[key] ?? (target[key] = {});
                mergeDeep(val, source[key]);
            } else {
                target[key] = source[key];
            }
        }
    }

    return mergeDeep(target, ...sources);
}

export function parseCssRules(text: string): Record<string, string> {
	const result: Record<string, string> = {};

	for (const rule of text.split(';')) {
		const [key, val] = rule.split(':');
		result[key] = val;
	}

	return result
}

export function formatCssRules(style: Record<string, string>): string {
	return Object.entries(style).map((k, v) => `${k}: ${v}`).join(';');
}

export function asArray<T>(val: T | T[]): T[] {
	return Array.isArray(val) ? val : [val];
}

export function clamp(val, min, max) {
    return min > val ? min : (max < val ? max : val);
}

/** Convert an EMF (Enhanced Metafile) ArrayBuffer to an SVG data-URL.
 *  Returns null if the buffer is not a valid EMF or cannot be converted. */
export function emfToSvgDataUrl(buffer: ArrayBuffer): string | null {
    try {
        const buf = new Uint8Array(buffer);
        const dv = new DataView(buffer);
        if (dv.getUint32(0, true) !== 1) return null; // not EMR_HEADER

        const headerSize = dv.getUint32(4, true);
        const boundsL = dv.getInt32(8, true), boundsT = dv.getInt32(12, true);
        const boundsR = dv.getInt32(16, true), boundsB = dv.getInt32(20, true);

        let winOrgX = 0, winOrgY = 0, winExtX = boundsR - boundsL, winExtY = boundsB - boundsT;
        let vpExtX = boundsR - boundsL, vpExtY = boundsB - boundsT;
        let curX = 0, curY = 0;
        let pathData = '';
        let inPath = false;
        const objects: Record<number, { type: string; color: string }> = {};
        let brushColor = '#000000';
        const svgPaths: string[] = [];

        let offset = headerSize;
        while (offset < buf.length - 8) {
            const type = dv.getUint32(offset, true);
            const size = dv.getUint32(offset + 4, true);
            if (size < 8 || offset + size > buf.length) break;
            if (type === 14) break; // EMR_EOF

            switch (type) {
                case 10: winOrgX = dv.getInt32(offset+8,true); winOrgY = dv.getInt32(offset+12,true); break;
                case 9:  winExtX = dv.getInt32(offset+8,true); winExtY = dv.getInt32(offset+12,true); break;
                case 11: vpExtX = dv.getInt32(offset+8,true); vpExtY = dv.getInt32(offset+12,true); break;
                case 27: // MOVETOEX
                    curX = dv.getInt32(offset+8,true); curY = dv.getInt32(offset+12,true);
                    if (inPath) pathData += `M${curX} ${curY}`;
                    break;
                case 89: { // POLYLINETO16
                    const cnt = dv.getUint32(offset+24,true); let p = offset+28;
                    for (let i = 0; i < cnt; i++, p += 4) {
                        const x = dv.getInt16(p,true), y = dv.getInt16(p+2,true);
                        if (inPath) pathData += `L${x} ${y}`;
                        curX = x; curY = y;
                    }
                    break;
                }
                case 88: { // POLYBEZIERTO16
                    const cnt = dv.getUint32(offset+24,true); let p = offset+28;
                    for (let i = 0; i+2 < cnt; i += 3, p += 12) {
                        const x1=dv.getInt16(p,true),y1=dv.getInt16(p+2,true),
                              x2=dv.getInt16(p+4,true),y2=dv.getInt16(p+6,true),
                              x3=dv.getInt16(p+8,true),y3=dv.getInt16(p+10,true);
                        if (inPath) pathData += `C${x1} ${y1} ${x2} ${y2} ${x3} ${y3}`;
                        curX = x3; curY = y3;
                    }
                    break;
                }
                case 59: inPath = true; pathData = ''; break; // BEGINPATH
                case 60: inPath = false; break; // ENDPATH
                case 61: if (inPath) pathData += 'Z'; break; // CLOSEFIGURE
                case 62: // FILLPATH
                case 63: // STROKEANDFILLPATH
                    if (pathData) svgPaths.push(`<path d="${pathData}" fill="${brushColor}" fill-rule="evenodd"/>`);
                    pathData = '';
                    break;
                case 64: // STROKEPATH
                    if (pathData) svgPaths.push(`<path d="${pathData}" fill="none" stroke="${brushColor}" stroke-width="1"/>`);
                    pathData = '';
                    break;
                case 39: { // CREATEBRUSHINDIRECT
                    const ih = dv.getUint32(offset+8,true), st = dv.getUint32(offset+12,true);
                    const r = buf[offset+16], g = buf[offset+17], b = buf[offset+18];
                    objects[ih] = { type: 'brush', color: st === 0 ? `rgb(${r},${g},${b})` : 'none' };
                    break;
                }
                case 37: { // SELECTOBJECT
                    const ih = dv.getUint32(offset+8,true);
                    if (ih & 0x80000000) {
                        const s = ih & 0x7FFFFFFF;
                        if (s === 0) brushColor = '#ffffff';
                        else if (s === 7) brushColor = '#000000';
                        else if (s === 5) brushColor = 'none';
                    } else if (objects[ih]?.type === 'brush') {
                        brushColor = objects[ih].color;
                    }
                    break;
                }
                case 40: delete objects[dv.getUint32(offset+8,true)]; break; // DELETEOBJECT
            }
            offset += size;
        }

        if (svgPaths.length === 0) return null;
        const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${winOrgX} ${winOrgY} ${winExtX} ${winExtY}" width="${vpExtX}" height="${vpExtY}">${svgPaths.join('')}</svg>`;
        return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
    } catch {
        return null;
    }
}