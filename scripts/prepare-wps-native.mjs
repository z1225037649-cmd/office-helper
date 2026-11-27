import { readFile, writeFile } from 'node:fs/promises';

export async function prepareWpsNative() {
  const token = (await readFile('.local/native-token', 'utf8')).trim();
  const check = process.argv.includes('--sheet-average') ? '&check=sheet-average' : '';
  await writeFile(
    '.local/wps-native-test/main.js',
    `(function(){
    var pane; var count = 0;
    function report(stage, detail) {
      var xhr = new XMLHttpRequest();
      xhr.open('POST', 'http://127.0.0.1:5173/__native-report');
      xhr.setRequestHeader('Content-Type','application/json');
      xhr.setRequestHeader('X-Helper-Test',${JSON.stringify(token)});
      xhr.send(JSON.stringify({host:'wps',stage:stage,detail:detail}));
    }
    report('bootstrap-loaded', navigator.userAgent);
    window.OnOpenHelper = function(){
      var api = window.wps;
      var app = window.Application || (api && api.Application) || api;
      var paneHost = app && app.CreateTaskPane ? app : api;
      if (!app || !paneHost) return;
      var kind = 'text'; var doc;
      try { doc = app.ActiveWorkbook; if(doc) kind = 'sheet'; } catch(e){}
      if(!doc) { try { doc = app.ActiveDocument; } catch(e){} }
      if(!doc || !/^OfficeHelper-Native-Test\\.(docx|xlsx)$/.test(doc.Name) || pane) return;
      var created = paneHost.CreateTaskPane('http://127.0.0.1:5173/native-check.html?host=wps&kind='+kind+'&token=${token}${check}');
      pane = typeof created === 'number' || typeof created === 'string' ? paneHost.GetTaskPane(created) : created;
      pane.Width=400; pane.Visible=true;
      report('pane-opened',kind);
    };
    function begin(){
      var timer=setInterval(function(){ count++;
        try{window.OnOpenHelper();}catch(e){if(count===1)report('bootstrap-error',String(e));}
        if(pane || count>90) clearInterval(timer);
      },700); return true;
    }
    window.OnAddinLoad = function(){report('addin-loaded',''); return true;};
    setTimeout(function(){report('bridge-detected',{app:typeof window.Application,wps:typeof window.wps}); begin();},1200);
  })();`,
  );
}

await prepareWpsNative();
