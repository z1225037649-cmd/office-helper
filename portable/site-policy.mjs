// Shared by the package preflight and the loopback server. A new public entry
// must be explicitly reviewed here before it can be shipped or served.
const entries = new Set(['index.html', 'favicon.svg']);
for (const component of ['', 'writer/', 'sheet/']) {
  for (const file of ['index.html', 'main.js', 'ribbon.xml', 'manifest.xml']) {
    entries.add(`wps/${component}${file}`);
  }
}

export function isPortableSitePath(path) {
  if (typeof path !== 'string' || path.split('/').some((part) => !part || part.startsWith('.')))
    return false;
  if (entries.has(path)) return true;
  return (
    /^assets\/[A-Za-z0-9][A-Za-z0-9_.-]*\.(?:js|css|svg|png|woff2)$/.test(path) &&
    !/(?:^|[._/-])(?:native-check|diagnostics?|debug|tests?|tokens?|private|secrets?|certificates?)(?:[._-]|$)/i.test(
      path,
    )
  );
}

export const requiredPortableSiteFiles = [...entries, 'assets/icon-32.png', 'assets/icon-64.png'];
