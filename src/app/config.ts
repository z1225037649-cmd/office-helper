export const appConfig = {
  name: '小帮手',
  featureFlags: {
    'text.enlarge': true,
    'text.bold': true,
    'paragraph.center': true,
    'paragraph.spacing': true,
    'paragraph.indent': true,
    'sheet.sum': true,
    'sheet.average': true,
    'selection.inspect': true,
  } satisfies Record<string, boolean>,
};
