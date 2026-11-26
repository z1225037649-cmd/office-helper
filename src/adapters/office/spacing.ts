/**
 * Conservative proof of *explicit* multiple line spacing in canonical Word OOXML.
 * This only gates a native property write; it never modifies or reinserts XML.
 * Unknown namespaces, style-inherited rules, tables and revision markup fail closed.
 */
export function explicitMultipleSpacing(xml: string, values: readonly number[]): boolean[] {
  const unsupported = values.map(() => false);
  if (/<!DOCTYPE|<!ENTITY/i.test(xml)) return unsupported;
  const document = xml.match(/<w:document\b[^>]*>([\s\S]*?)<\/w:document>/)?.[1];
  const body = document?.match(/<w:body\b[^>]*>([\s\S]*?)<\/w:body>/)?.[1];
  if (!body || /<w:(?:tbl|sdt|ins|del|pPrChange|moveFrom|moveTo)\b/.test(body)) return unsupported;
  const paragraphs = [...body.matchAll(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g)];
  if (paragraphs.length !== values.length) return unsupported;
  return paragraphs.map(([p], i) => {
    const properties = p.match(/^<w:p\b[^>]*>\s*<w:pPr\b[^>]*>([\s\S]*?)<\/w:pPr>/)?.[1];
    const spacing = properties?.match(/<w:spacing\b[^>]*\/>/g);
    if (spacing?.length !== 1) return false;
    const rule = spacing[0]!.match(/\bw:lineRule="([^"]*)"/)?.[1];
    const line = spacing[0]!.match(/\bw:line="(\d+)"/)?.[1];
    return (
      rule === 'auto' &&
      !!line &&
      Number(line) > 0 &&
      Math.abs(Number(line) / 20 - values[i]!) < 0.001
    );
  });
}
