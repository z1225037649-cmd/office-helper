export const appConfig = {
  name: '小帮手',
  featureFlags: {
    'text.enlarge': true,
    'sheet.sum': true,
    'selection.inspect': true,
  } satisfies Record<string, boolean>,
};
