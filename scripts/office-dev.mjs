import { mkdir, copyFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { generateSideloadFile, registerAddIn, unregisterAddIn } from 'office-addin-dev-settings';
import { OfficeAddinManifest, OfficeApp } from 'office-addin-manifest';

const manifestPath = resolve('manifests/office/manifest.xml');
const action = process.argv[2] ?? 'install';
if (action === 'remove') {
  await unregisterAddIn(manifestPath);
  console.log('已移除办公小帮手的 Office 开发注册。');
} else if (action === 'install') {
  await registerAddIn(manifestPath);
  const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);
  await mkdir('output/samples', { recursive: true });
  for (const [app, extension] of [
    [OfficeApp.Word, 'docx'],
    [OfficeApp.Excel, 'xlsx'],
  ]) {
    const generated = await generateSideloadFile(app, manifest);
    const destination = `output/samples/OfficeHelper-Sample.${extension}`;
    // 已填入内容的试用文件保留，避免重复安装覆盖用户试用结果。
    try {
      await copyFile(generated, destination, 1);
    } catch (error) {
      if (error.code !== 'EEXIST') throw error;
    }
  }
  console.log('已注册正常插件入口，试用文件位于 output/samples。请先启动 npm run dev:office。');
} else {
  throw new Error('支持 install 或 remove。');
}
