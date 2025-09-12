import http from 'http';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = 8080;
const PLUGIN_DIR = path.join(__dirname, 'public/plugins/insert-hello');

const mimeTypes = {
  '.html': 'text/html',
  '.js': 'text/javascript',
  '.json': 'application/json',
  '.png': 'image/png',
  '.ico': 'image/x-icon'
};

const server = http.createServer((req, res) => {
  console.log(`${new Date().toISOString()} - ${req.method} ${req.url}`);
  
  // 设置CORS头
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
  if (req.method === 'OPTIONS') {
    res.writeHead(200);
    res.end();
    return;
  }
  
  let filePath;
  if (req.url === '/config.json') {
    filePath = path.join(PLUGIN_DIR, 'config.json');
  } else if (req.url === '/index.html') {
    filePath = path.join(PLUGIN_DIR, 'index.html');
  } else if (req.url === '/plugin.js') {
    filePath = path.join(PLUGIN_DIR, 'plugin.js');
  } else if (req.url === '/icon.png') {
    filePath = path.join(PLUGIN_DIR, 'icon.png');
  } else {
    res.writeHead(404);
    res.end('Not Found');
    return;
  }
  
  fs.readFile(filePath, (err, data) => {
    if (err) {
      console.error('文件读取错误:', err);
      res.writeHead(404);
      res.end('File not found');
      return;
    }
    
    const ext = path.extname(filePath);
    const mimeType = mimeTypes[ext] || 'application/octet-stream';
    
    res.writeHead(200, { 'Content-Type': mimeType });
    res.end(data);
  });
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`插件服务器启动在: http://0.0.0.0:${PORT}`);
  console.log(`本地访问: http://localhost:${PORT}/config.json`);
  console.log(`局域网访问: http://192.168.0.10:${PORT}/config.json`);
  console.log(`Docker访问: http://172.17.0.1:${PORT}/config.json`);
});