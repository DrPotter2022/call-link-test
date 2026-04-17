const http = require('http');
const fs = require('fs');
const path = require('path');
const root = process.cwd();
const mime = { '.html':'text/html; charset=utf-8', '.json':'application/json; charset=utf-8', '.kml':'application/vnd.google-earth.kml+xml; charset=utf-8', '.js':'text/javascript; charset=utf-8', '.css':'text/css; charset=utf-8' };
http.createServer((req,res)=>{
  let p = decodeURIComponent((req.url||'/').split('?')[0]);
  if (p === '/' || p === '') p = '/index.html';
  const f = path.normalize(path.join(root, p));
  if (!f.startsWith(root)) { res.statusCode=403; return res.end('Forbidden'); }
  fs.readFile(f,(e,d)=>{
    if(e){res.statusCode=404;return res.end('Not Found');}
    res.statusCode=200;
    res.setHeader('Content-Type', mime[path.extname(f).toLowerCase()] || 'application/octet-stream');
    res.end(d);
  });
}).listen(8080,'127.0.0.1');
