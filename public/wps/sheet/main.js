/* WPS 功能区运行时与任务窗格运行时彼此独立。入口只负责打开共用界面。 */
(function () {
  'use strict';
  var pane;
  var base = window.location.origin + '/';

  window.OnAddinLoad = function () {
    return true;
  };
  window.OnOpenHelper = function () {
    var api = window.wps;
    var app = window.Application || (api && api.Application) || api;
    var paneHost = app && app.CreateTaskPane ? app : api;
    if (!app || !paneHost || typeof paneHost.CreateTaskPane !== 'function') {
      window.alert('请从支持加载项的 WPS 内打开小帮手。');
      return;
    }
    if (pane) {
      try {
        pane.Visible = true;
        return;
      } catch (_) {
        pane = undefined;
      }
    }
    var kind = 'text';
    try {
      if (app.ActiveWorkbook) kind = 'sheet';
    } catch (_) {
      /* 文字宿主无 ActiveWorkbook。 */
    }
    var created = paneHost.CreateTaskPane(base + '?host=wps&kind=' + kind);
    pane =
      typeof created === 'number' || typeof created === 'string'
        ? paneHost.GetTaskPane(created)
        : created;
    pane.Width = 400;
    pane.Visible = true;
  };
})();
