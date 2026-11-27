// Build-time helpers only; deliberately excluded from the distributed app/.
import { createHash } from 'node:crypto';
import { lstat, mkdir, readFile, readdir, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { isPortableSitePath, requiredPortableSiteFiles } from './site-policy.mjs';

export function portablePackagePaths(output, version) {
  if (
    typeof version !== 'string' ||
    !/^(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)$/.test(version) ||
    version.split('.').some((part) => Number(part) > 65535)
  ) {
    throw new Error('Portable release version must be three integers between 0 and 65535.');
  }
  const name = `OfficeHelper-${version}-win-x64`;
  const destination = join(resolve(output), name);
  return {
    name,
    destination,
    archive: destination + '.zip',
    checksum: destination + '.zip.sha256',
  };
}

export async function assertPortableOutputsAbsent(paths, { includeDirectory = true } = {}) {
  const outputs = includeDirectory
    ? [paths.destination, paths.archive, paths.checksum]
    : [paths.archive, paths.checksum];
  for (const path of outputs) {
    try {
      await lstat(path);
    } catch (error) {
      if (error.code === 'ENOENT') continue;
      throw error;
    }
    throw new Error(
      `Portable output already exists; preserve it and choose a new version: ${path}`,
    );
  }
}

export function assertSafePortableContent(bytes, label) {
  for (const encoding of ['utf8', 'utf16le']) {
    const text = Buffer.from(bytes).toString(encoding);
    if (
      /native-check\.html|__native-report|x-helper-test|OFFICE_HELPER_NATIVE_TEST|native-token|sourceMappingURL\s*=/i.test(
        text,
      ) ||
      /-----BEGIN (?:[A-Z0-9]+ )*PRIVATE KEY-----/i.test(text) ||
      /(?:^|[^a-f0-9])[a-f0-9]{48}(?![a-f0-9])/i.test(text) ||
      /[?&]token\s*=/i.test(text)
    ) {
      // Never echo the matched data: a diagnostic token or private key may be present.
      throw new Error(
        `Diagnostics, credentials or source maps are forbidden in portable file: ${label}`,
      );
    }
  }
}

const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');

export async function inspectPortableSite(root) {
  const files = [];
  const allowedDirectories = new Set(['', 'assets', 'wps', 'wps/writer', 'wps/sheet']);
  async function visit(relative) {
    const path = join(root, relative);
    const info = await lstat(path);
    if (info.isSymbolicLink())
      throw new Error(`Portable site cannot contain links: ${relative || '.'}`);
    if (info.isDirectory()) {
      if (!allowedDirectories.has(relative))
        throw new Error(`Unexpected portable site directory: ${relative}`);
      const children = await readdir(path);
      for (const child of children.sort()) await visit(relative ? `${relative}/${child}` : child);
      return;
    }
    if (!info.isFile() || !isPortableSitePath(relative))
      throw new Error(`Unexpected portable site file: ${relative}`);
    const bytes = await readFile(path);
    assertSafePortableContent(bytes, relative);
    files.push({ path: relative, sha256: digest(bytes) });
  }
  await visit('');
  for (const path of requiredPortableSiteFiles) {
    if (!files.some((file) => file.path === path))
      throw new Error(`Missing portable site entry: ${path}`);
  }
  if (!files.some((file) => /^assets\/.+\.js$/.test(file.path)))
    throw new Error('The portable site has no compiled JavaScript bundle.');
  return files;
}

export async function copyInspectedPortableSite(source, destination, files) {
  for (const file of files) {
    const sourcePath = join(source, file.path);
    const info = await lstat(sourcePath);
    if (!info.isFile() || info.isSymbolicLink())
      throw new Error(`Portable site changed after preflight: ${file.path}`);
    const bytes = await readFile(sourcePath);
    if (digest(bytes) !== file.sha256)
      throw new Error(`Portable site changed after preflight: ${file.path}`);
    const target = join(destination, file.path);
    await mkdir(dirname(target), { recursive: true });
    await writeFile(target, bytes, { flag: 'wx' });
  }
}

export function createPortableManifest(original, version) {
  portablePackagePaths('.', version);
  if (
    (original.match(/<Version>[^<]+<\/Version>/g) ?? []).length !== 1 ||
    (original.match(/<Id>26b5f5ec-853a-4c06-9c1c-8ac9ea33421a<\/Id>/g) ?? []).length !== 1
  ) {
    throw new Error('The portable manifest needs exactly one development ID and version.');
  }
  const manifest = original
    .replace(/<Version>[^<]+<\/Version>/, `<Version>1.${version}</Version>`)
    .replaceAll('26b5f5ec-853a-4c06-9c1c-8ac9ea33421a', 'f5202c69-1f59-4e09-b510-829a3748a31b')
    .replaceAll('localhost:3443', 'localhost:18766')
    .replace('办公小帮手（开发版）', '办公小帮手（便携试用版）')
    .replace(
      /(<Description\b[^>]*\bDefaultValue=")[^"]*(")/,
      '$1预览确认文字和表格排版；合计、平均值及重复漏填检查；打印设置、副本与PDF、独立通知模板。可用能力以当前宿主为准。$2',
    );
  assertSafePortableContent(Buffer.from(manifest), 'manifests/office.xml');
  if (/(?:localhost|127\.0\.0\.1):(?:3443|5173)\b|办公小帮手（开发版）/.test(manifest))
    throw new Error('Development settings remain in the portable manifest.');
  return manifest;
}

const maskComments = (xml) => xml.replace(/<!--[\s\S]*?-->/g, (match) => ' '.repeat(match.length));

export function populateWorksheetSample(xml) {
  const matches = [
    ...maskComments(xml).matchAll(
      /<((?:[A-Za-z_][\w.-]*:)?sheetData)\b([^>]*?)(?:\/>|>[\s\S]*?<\/\1\s*>)/g,
    ),
  ];
  if (matches.length !== 1)
    throw new Error('The Microsoft Excel template must contain exactly one supported sheetData.');
  const [match] = matches;
  const [, tag, attributes] = match;
  const prefix = tag.slice(0, -'sheetData'.length);
  const textCell = (address, text) =>
    `<${prefix}c r="${address}" t="inlineStr"><${prefix}is><${prefix}t>${text}</${prefix}t></${prefix}is></${prefix}c>`;
  const rows =
    `<${prefix}row r="1">${textCell('A1', '项目')}${textCell('B1', '金额')}</${prefix}row>` +
    [120, 85, 45.5, 200]
      .map(
        (value, i) =>
          `<${prefix}row r="${i + 2}">${textCell(`A${i + 2}`, `项目${i + 1}`)}<${prefix}c r="B${i + 2}"><${prefix}v>${value}</${prefix}v></${prefix}c></${prefix}row>`,
      )
      .join('');
  let updated =
    xml.slice(0, match.index) +
    `<${tag}${attributes}>${rows}</${tag}>` +
    xml.slice(match.index + match[0].length);
  const dimension = /<((?:[A-Za-z_][\w.-]*:)?dimension)\b([^>]*?)(?:\/>|>[\s\S]*?<\/\1\s*>)/.exec(
    maskComments(updated),
  );
  if (dimension) {
    const attributes = dimension[2].replace(/\s+ref\s*=\s*(?:"[^"]*"|'[^']*')/g, '');
    updated =
      updated.slice(0, dimension.index) +
      `<${dimension[1]}${attributes} ref="A1:B5"/>` +
      updated.slice(dimension.index + dimension[0].length);
  }
  return updated;
}
