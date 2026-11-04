/** Convert an EMF (Enhanced Metafile) ArrayBuffer to an SVG data-URL.
 * Returns null if the buffer is not a valid EMF or cannot be converted. */
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
