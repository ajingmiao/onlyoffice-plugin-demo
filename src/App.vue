<template>
  <div style="height:100vh;display:flex;flex-direction:column;">
    <header style="padding:8px;border-bottom:1px solid #eee;display:flex;gap:8px;align-items:center;">
      <strong>ONLYOFFICE Vue3 Demo</strong>
      <button @click="reload" style="margin-left:auto;">重载编辑器</button>
    </header>
    <div id="editor" style="flex:1;"></div>
  </div>
</template>

<script setup lang="ts">
import { onMounted, onBeforeUnmount } from 'vue'

let editor: any = null

// 你的 DocumentServer 根地址（不含 /web-apps/...）
const DOCUMENT_SERVER = 'http://localhost:9998/' // ← 修改这里

// 当前页面可访问的 docx（放在 public/ 目录）
const FILE_URL = '/test.docx'

// 插件配置文件 URL（编辑器将从这里直接加载）
const PLUGIN_CONFIG_URL = '/plugins/insert-hello/config.json'

function createEditor() {
  // @ts-ignore
  const DocsAPI = (window as any).DocsAPI
  if (!DocsAPI) {
    alert('DocsAPI 未加载，请检查 index.html 中的 <script src="...api.js"> 地址是否正确')
    return
  }
const FILE_URL = 'https://temp2-1302420147.cos.ap-nanjing.myqcloud.com/test.docx'

const cfg = {
  document: {
    fileType: 'docx',
    title: 'Demo Template.docx',
    url: FILE_URL, // 远端可访问的地址
  },
  documentType: 'word',
  editorConfig: {
    lang: 'zh-CN',
    customization: { plugins: true },
  },
  height: '100%',
  width: '100%',
  plugins: {
    pluginsData: [
      location.origin + '/plugins/insert-hello/config.json' // 插件还是本地提供
    ]
  }
}

  // 销毁旧实例
  if (editor && editor.destroyEditor) editor.destroyEditor()

  // 创建实例
  // @ts-ignore
  editor = new DocsAPI.DocEditor('editor', cfg)

  // 事件：文档就绪
  // @ts-ignore
  editor.events && editor.events.on('onDocumentReady', () => {
    console.log('[onDocumentReady] 文档可用')
  })
}

function reload() {
  createEditor()
}

onMounted(() => createEditor())
onBeforeUnmount(() => {
  if (editor && editor.destroyEditor) editor.destroyEditor()
})
</script>
