<template>
  <div style="height:100vh;display:flex;flex-direction:column;">
    <header style="padding:8px;border-bottom:1px solid #eee;display:flex;gap:8px;align-items:center;">
      <strong>ONLYOFFICE Vue3 Demo</strong>
      <button @click="insertPlaceholder" style="background:#007bff;color:#fff;border:none;padding:6px 12px;border-radius:4px;">插入 {{ name }}</button>
      <button @click="insertTable" style="background:#28a745;color:#fff;border:none;padding:6px 12px;border-radius:4px;">插入表格</button>
      <button @click="reload" style="margin-left:auto;">重载编辑器</button>
    </header>
    <div id="editor" style="flex:1;"></div>
  </div>
</template>

<script setup lang="ts">
import { onMounted, onBeforeUnmount, ref } from 'vue'

let editor: any = null
const name = ref('name')

// —— 基本配置（按你的环境修改）——
const DOCUMENT_SERVER = 'http://localhost:9998'                     // DocumentServer 根地址
const DOCS_ORIGIN = new URL(DOCUMENT_SERVER).origin                 // 必须与 iframe.src 的 origin 完全一致
const FILE_URL = 'https://temp2-1302420147.cos.ap-nanjing.myqcloud.com/test.docx'

// 与插件一致（仅用于自动加载，不用于消息通道）
const PLUGIN_GUID = 'asc.{6A1D2E30-1B7D-4A87-A2D7-4F3BB8A3C9E1}'
// 一定要能在浏览器直接 200 打开（建议用带版本前缀的路径）
const PLUGIN_CONF = 'http://localhost:9998/8.3.3-c7acbc51525f4e8ddd251a711f871b28/sdkjs-plugins/demo-a/config.json'


fetch(PLUGIN_CONF)
  .then(r => r.text())
  .then(t => { console.log('原文:', t); console.log('能否JSON.parse:', !!JSON.parse(t)); });

// —— 获取编辑器 iframe —— 
function getEditorIframe(): HTMLIFrameElement | null {
  const iframes = Array.from(document.querySelectorAll('iframe')) as HTMLIFrameElement[]
  for (const f of iframes) {
    try {
      if (new URL(f.src).origin === DOCS_ORIGIN && f.contentWindow) return f
    } catch {}
  }
  return null
}
async function waitForIframe(retries = 40, interval = 150): Promise<HTMLIFrameElement | null> {
  for (let i = 0; i < retries; i++) {
    const f = getEditorIframe()
    if (f?.contentWindow) return f
    await new Promise(r => setTimeout(r, interval))
  }
  return null
}

// —— 宿主 → 插件（普通 postMessage；插件侧监听 window.addEventListener('message')）——
async function sendToPluginBridge(cmd: string, payload: any) {
  const iframe = await waitForIframe()
  if (!iframe) {
    alert('未找到编辑器 iframe（可能尚未加载完成）')
    return
  }
  const msg = { type: 'bridgeCommand', cmd, payload }
  // targetOrigin 必须与 iframe.src 的 origin 完全一致（此处为 http://localhost:9998）
  iframe.contentWindow!.postMessage(msg, DOCS_ORIGIN)
}

// —— 创建/重载编辑器 —— 
function createEditor() {
  // @ts-ignore
  const DocsAPI = (window as any).DocsAPI
  if (!DocsAPI) {
    alert('DocsAPI 未加载：请检查 <script src=".../web-apps/apps/api/documents/api.js">')
    return
  }

  // 销毁旧实例
  if (editor?.destroyEditor) editor.destroyEditor()

  const cfg = {
    document: {
      fileType: 'docx',
      title: 'Demo Template.docx',
      url: FILE_URL,
      key: 'test-1',
      permissions: { edit: true, review: true, comment: true }   // ✅ 确保可编辑
    },
    documentType: 'word',
    editorConfig: {
      lang: 'zh-CN',
      mode: 'edit',                                               // ✅ 编辑模式
      user: { id: 'u1', name: 'Tester' },
      customization: { plugins: true },
      plugins: {
        autostart: [PLUGIN_GUID],                                 // 自动加载你的插件
        pluginsData: [{ url: '/8.3.3-*/sdkjs-plugins/demo-a/config.json?v=20250911' }]                       // 指向插件 config.json
      }
    },
    width: '100%',
    height: '100%'
  }

  // @ts-ignore
  editor = new DocsAPI.DocEditor('editor', cfg)
}

function reload() { createEditor() }

// —— 业务按钮 —— 
async function insertPlaceholder() {
  await sendToPluginBridge('insertText', { text: `{{${name.value}}}` })
}
async function insertTable() {
  await sendToPluginBridge('insertTable', {
    rows: 3, cols: 3,
    headers: ['A','B','C'],
    data: [['1','2','3'], ['4','5','6'], ['7','8','9']]
  })
}

// —— 生命周期 —— 
onMounted(() => createEditor())
onBeforeUnmount(() => { if (editor?.destroyEditor) editor.destroyEditor() })
</script>
