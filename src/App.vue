<template>
  <div style="height:100vh;display:flex;flex-direction:column;">
    <header style="padding:8px;border-bottom:1px solid #eee;display:flex;gap:8px;align-items:center;">
      <strong>ONLYOFFICE Vue3 Demo</strong>
      <button @click="insertText" style="background:#007bff;color:#fff;border:none;padding:6px 12px;border-radius:4px;">插入文字</button>
      <button @click="insertLink" style="background:#1e88e5;color:#fff;border:none;padding:6px 12px;border-radius:4px;">
        插入链接
      </button>
      <span :style="{marginLeft:'auto', color: pluginReady ? '#28a745' : '#999'}">
        {{ pluginReady ? '插件已就绪' : '等待插件就绪…' }}
      </span>
    </header>
    <div id="editor" style="flex:1;"></div>
  </div>
</template>

<script setup lang="ts">
import { onMounted, onBeforeUnmount, ref } from 'vue'

let editor: any = null
const pluginReady = ref(false)

// === 按你的环境修改 ===
const DOCUMENT_SERVER = 'http://192.168.1.103:9998'
//const DOCUMENT_SERVER = 'http://localhost:9998'
const DOCS_API = DOCUMENT_SERVER + '/web-apps/apps/api/documents/api.js'
const FILE_URL = 'https://temp2-1302420147.cos.ap-nanjing.myqcloud.com/test.docx'

// 插件 GUID（需与插件 config.json 一致）  
//9.0.4-9ade76efaf7465c8db6be392804370a8

// http://192.168.1.103:9998/9.0.4-870a82e9bc8b96c4c53877e589326856/web-apps/apps/documenteditor/main/index.html?_dc=9.0.4-50&lang=zh-CN&customer=ONLYOFFICE&type=desktop&frameEditorId=editor&isForm=false&parentOrigin=http://localhost:5174&fileType=docx

const PLUGIN_GUID = 'asc.{6A1D2E30-1B7D-4A87-A2D7-4F3BB8A3C9E1}'
// 插件 config.json（必须可 200 直连）
const PLUGIN_CONF = DOCUMENT_SERVER + '/9.0.4-870a82e9bc8b96c4c53877e589326856/sdkjs-plugins/demo-a/config.json'

// ---- 动态加载 api.js ----
function loadDocsApi(): Promise<void> {
  return new Promise((resolve, reject) => {
    if ((window as any).DocsAPI) return resolve()
    const script = document.createElement('script')
    script.src = DOCS_API
    script.onload = () => resolve()
    script.onerror = () => reject(new Error('加载 DocsAPI 失败：' + DOCS_API))
    document.head.appendChild(script)
  })
}

// ---- 命令发送：带队列 + 重试，确保等插件就绪 ----
type Cmd = { cmd: string; payload: any }
const queue: Cmd[] = []
let flushTimer: number | null = null

function serviceCommandSafe(cmd: string, payload: any) {
  // 未加载好编辑器
  if (!editor || typeof editor.serviceCommand !== 'function') {
    console.warn('[ONLYOFFICE] editor/serviceCommand 未就绪，先入队：', cmd, payload)
    queue.push({ cmd, payload })
    return
  }
  // 插件未就绪：排队等待
  if (!pluginReady.value) {
    console.log('[ONLYOFFICE] 插件未就绪，先入队：', cmd, payload)
    queue.push({ cmd, payload })
    scheduleFlush()
    return
  }
  try {
      if (typeof editor.executeCommand === 'function') {
        editor.executeCommand('focus');
      }
      // 让浏览器完成一次光标同步，然后再发命令
      setTimeout(() => editor.serviceCommand(cmd, payload), 0);
  } catch (e) {
    console.error('[ONLYOFFICE] serviceCommand 调用失败：', e)
  }
}

function scheduleFlush() {
  if (flushTimer !== null) return
  flushTimer = window.setInterval(() => {
    if (!editor || typeof editor.serviceCommand !== 'function') return
    if (!pluginReady.value) return
    // 开始出队
    while (queue.length) {
      const { cmd, payload } = queue.shift()!
      try {
        console.log('[ONLYOFFICE] 发送排队指令：', cmd, payload)
        editor.serviceCommand(cmd, payload)
      } catch (e) {
        console.error('[ONLYOFFICE] 队列指令发送失败：', cmd, e)
      }
    }
    if (flushTimer !== null) {
      clearInterval(flushTimer)
      flushTimer = null
    }
  }, 300)
}

// ---- 创建编辑器 ----
async function createEditor() {
  await loadDocsApi()
  // @ts-ignore
  const DocsAPI = (window as any).DocsAPI

  if (editor?.destroyEditor) editor.destroyEditor()
  pluginReady.value = false

  const cfg = {
    document: {
      fileType: 'docx',
      title: 'Demo.docx',
      url: FILE_URL,
      key: 'test-key',
      permissions: { edit: true }
    },
    documentType: 'word',
    editorConfig: {
      lang: 'zh-CN',
      mode: 'edit',
      user: { id: 'u1', name: 'Tester' }
    },
    plugins: {
      autostart: [PLUGIN_GUID],
      pluginsData: [{ url: PLUGIN_CONF + '?v=' + Date.now() }] // 强刷缓存
    },
    width: '100%',
    height: '100%',
    events: {
      onPluginsReady: () => {
        console.log('[ONLYOFFICE] 插件资源已加载，但不代表已初始化');
      },
      onDocumentReady: () => {
        console.log('[ONLYOFFICE] 文档已就绪')
        // 有些插件会很快 ready，这里安排一次队列尝试
        autoOpenPluginSidePanel();
        scheduleFlush();
      },
      onInfo: (e: any) => {
         const msg = e?.data;
        console.log('[ONLYOFFICE] Info 事件：', e.data)
        if (!e) return
        if (msg?.command === 'openPlugin') {
           console.log('[ONLYOFFICE] openPlugin')
          // 这里再把消息转发给 DS 内部，让它真的打开 UI 面板
          editor.frame?.contentWindow?.Common?.Gateway?.sendInfo(msg);
        }
        if (e.data.command === 'pluginInitialized') {
          pluginReady.value = true
          console.log('[ONLYOFFICE] 插件初始化完成')
          scheduleFlush()
        }
        if (e.data .command === 'pluginAck') {
          console.log('[ONLYOFFICE] 插件确认收到:', e.data)
        }
        if (e.data.command === 'pluginError') {
          console.error('[ONLYOFFICE] 插件错误:', e.data)
        }
      }
    }
  }

  editor = new DocsAPI.DocEditor('editor', cfg)
}

function autoOpenPluginSidePanel() {
  try {
    editor.executeCommand && editor.executeCommand('focus');

    // 通过你现有的 Gateway 让 WebApp 打开插件（指令名按你那边实现）
    editor.frame?.contentWindow?.Common?.Gateway?.sendInfo?.({
      command: 'openPlugin',
      data: {
        guid: 'asc.{6A1D2E30-1B7D-4A87-A2D7-4F3BB8A3C9E1}',
        variation: 'side panel'
      }
    });
  } catch(e) { console.error('auto-open plugin failed:', e); }
}

// ---- 按钮动作（发指令到插件）----
function insertText() {
  serviceCommandSafe('insertText', '【外部插入的文字】')
}
function insertLink() {
  // 例1：空链接占位（后续再填 URL）
  serviceCommandSafe('insertLink', {
    text: '点击这里',   // 显示文字
    url: '',            // 允许为空；插件端会用 about:blank 占位+打Tag
    bold: true,         // 加粗
    underline: false,    // 建议加下划线
    color: '#1976D2',    // 蓝色（可换）
    json:{name:'gm'}
  });
}


onMounted(() => createEditor())
onBeforeUnmount(() => {
  if (editor?.destroyEditor) editor.destroyEditor()
  if (flushTimer !== null) clearInterval(flushTimer)
})
</script>
