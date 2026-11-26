import { mkdir, copyFile, writeFile } from 'node:fs/promises';
for (const [host, name] of [
  ['writer', 'OfficeHelperWriter'],
  ['sheet', 'OfficeHelperSheet'],
]) {
  const destination = `public/wps/${host}`;
  await mkdir(destination, { recursive: true });
  for (const file of ['index.html', 'main.js', 'ribbon.xml'])
    await copyFile(`public/wps/${file}`, `${destination}/${file}`);
  await writeFile(
    `${destination}/manifest.xml`,
    `<?xml version="1.0" encoding="UTF-8"?><JsPlugin><ApiVersion>1.0.0</ApiVersion><Name>${name}</Name><Description>办公小帮手</Description></JsPlugin>`,
  );
}
