<template>
  <div style="height:100vh;display:flex;flex-direction:column;">
    <header style="padding:8px;border-bottom:1px solid #eee;display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
      <strong>ONLYOFFICE Vue3 Demo</strong>
      <button @click="insertText" style="background:#007bff;color:#fff;border:none;padding:6px 12px;border-radius:4px;">插入文字</button>
      <button @click="insertLink" style="background:#1e88e5;color:#fff;border:none;padding:6px 12px;border-radius:4px;">
        插入链接
      </button>

      <!-- WordArt 按钮组 -->
      <div style="display:flex;gap:4px;border-left:1px solid #ddd;padding-left:8px;">
        <button @click="insertCustomWordArt" style="background:#4caf50;color:#fff;border:none;padding:6px 12px;border-radius:4px;">
          自定义WordArt
        </button>
        <button @click="insertClassicWordArt" style="background:#ff9800;color:#fff;border:none;padding:6px 8px;border-radius:4px;">
          经典
        </button>
        <button @click="insertModernWordArt" style="background:#673ab7;color:#fff;border:none;padding:6px 8px;border-radius:4px;">
          现代
        </button>
        <button @click="insertFunWordArt" style="background:#e91e63;color:#fff;border:none;padding:6px 8px;border-radius:4px;">
          趣味
        </button>
      </div>

      <!-- Shape 测试按钮 -->
      <div style="display:flex;gap:4px;border-left:1px solid #ddd;padding-left:8px;">
        <button @click="testShapeInline" style="background:#795548;color:#fff;border:none;padding:6px 8px;border-radius:4px;font-size:12px;">
          测试Shape
        </button>
        <button @click="insertRectangleInline" style="background:#607d8b;color:#fff;border:none;padding:6px 8px;border-radius:4px;">
          矩形
        </button>
        <button @click="insertCircleInline" style="background:#455a64;color:#fff;border:none;padding:6px 8px;border-radius:4px;">
          圆形
        </button>
        <button @click="insertTriangleInline" style="background:#37474f;color:#fff;border:none;padding:6px 8px;border-radius:4px;">
          三角形
        </button>
      </div>

      <!-- Table 按钮组 -->
      <div style="display:flex;gap:4px;border-left:1px solid #ddd;padding-left:8px;">
        <button @click="insertSimpleTable" style="background:#8bc34a;color:#fff;border:none;padding:6px 8px;border-radius:4px;">
          简单表格
        </button>
        <button @click="insertScheduleTable" style="background:#4caf50;color:#fff;border:none;padding:6px 8px;border-radius:4px;">
          课程表
        </button>
        <button @click="insertComparisonTable" style="background:#66bb6a;color:#fff;border:none;padding:6px 8px;border-radius:4px;">
          对比表
        </button>
        <button @click="insertCustomTable" style="background:#388e3c;color:#fff;border:none;padding:6px 8px;border-radius:4px;">
          自定义
        </button>
        <button @click="insertDynamicTable" style="background:#2e7d32;color:#fff;border:none;padding:6px 8px;border-radius:4px;">
          动态数据
        </button>
      </div>

      <!-- Selection Binding 按钮组 -->
      <div style="display:flex;gap:4px;border-left:1px solid #ddd;padding-left:8px;">
        <button @click="analyzeSelection" style="background:#6a1b9a;color:#fff;border:none;padding:6px 8px;border-radius:4px;">
          分析选中
        </button>
        <button @click="bindAsTextField" style="background:#7b1fa2;color:#fff;border:none;padding:6px 8px;border-radius:4px;">
          绑定文本
        </button>
        <button @click="bindAsVariable" style="background:#8e24aa;color:#fff;border:none;padding:6px 8px;border-radius:4px;">
          绑定变量
        </button>
        <button @click="bindAsDataSource" style="background:#9c27b0;color:#fff;border:none;padding:6px 8px;border-radius:4px;">
          绑定数据源
        </button>
      </div>

      <!-- Chart Binding 按钮组 - 新增 -->
      <div style="display:flex;gap:4px;border-left:1px solid #ddd;padding-left:8px;">
        <button @click="bindChartData" style="background:#e65100;color:#fff;border:none;padding:6px 8px;border-radius:4px;">
          绑定图表数据
        </button>
        <button @click="getChartSummary" style="background:#ff6f00;color:#fff;border:none;padding:6px 8px;border-radius:4px;">
          图表摘要
        </button>
        <button @click="testChartDetection" style="background:#ff8f00;color:#fff;border:none;padding:6px 8px;border-radius:4px;">
          测试检测
        </button>
      </div>

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
const DOCUMENT_SERVER = 'http://localhost:9998'
//const DOCUMENT_SERVER = 'http://222.187.11.98:8918'
//const DOCUMENT_SERVER = 'http://localhost:9998'
const DOCS_API = DOCUMENT_SERVER + '/web-apps/apps/api/documents/api.js'
const FILE_URL = 'https://temp2-1302420147.cos.ap-nanjing.myqcloud.com/test.docx'

// 插件 GUID（需与插件 config.json 一致）  
//9.0.4-9ade76efaf7465c8db6be392804370a8
//http://222.187.11.98:8918/8.3.3-5dd8d105cac84554276dfe74dce59789/web-apps/apps/documenteditor/main/index.html?_dc=8.3.3-18&lang=zh-CN&customer=ONLYOFFICE&type=desktop&frameEditorId=editor&isForm=false&parentOrigin=http://localhost:5174&fileType=docx
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
      pluginsData: [{ url: PLUGIN_CONF + '?v=2.0.' + Date.now() }] // 强刷缓存
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
        if (e.data.command === 'linkClicked') {
          console.log('[ONLYOFFICE] 链接被点击，数据:', e.data.data)
          handleLinkClick(e.data.data)
        }
        if (e.data.command === 'wordArtInserted') {
          console.log('[ONLYOFFICE] WordArt 插入成功:', e.data.data)
          handleWordArtInserted(e.data.data)
        }
        if (e.data.command === 'wordArtError') {
          console.error('[ONLYOFFICE] WordArt 插入失败:', e.data.data)
          handleWordArtError(e.data.data)
        }
        if (e.data.command === 'tableClicked') {
          console.log('[ONLYOFFICE] 表格被点击，数据:', e.data.data)
          handleTableClick(e.data.data)
        }
        if (e.data.command === 'selectionAnalyzed') {
          console.log('[ONLYOFFICE] 选中内容分析结果:', e.data.data)
          handleSelectionAnalyzed(e.data.data)
        }
        if (e.data.command === 'selectionBound') {
          console.log('[ONLYOFFICE] 选中内容绑定完成:', e.data.data)
          handleSelectionBound(e.data.data)
        }
        if (e.data.command === 'bindingClicked') {
          console.log('[ONLYOFFICE] 绑定控件被点击:', e.data.data)
          handleBindingClicked(e.data.data)
        }
        // 通用元素检测事件 - 新增
        if (e.data.command === 'elementClicked') {
          console.log('[ONLYOFFICE] 元素被点击:', e.data.data)
          handleElementClicked(e.data.data)
        }
        // 精确表格单元格检测事件 - 新增
        if (e.data.command === 'preciseTableCellClicked') {
          console.log('[ONLYOFFICE] 精确表格单元格被点击:', e.data.data)
          handlePreciseTableCellClicked(e.data.data)
        }
        // 图表点击检测事件 - 新增
        if (e.data.command === 'chartClicked') {
          console.log('[ONLYOFFICE] 图表被点击:', e.data.data)
          handleChartClicked(e.data.data)
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

// ---- WordArt 相关功能 ----
function insertCustomWordArt() {
  // 自定义 WordArt 参数
  serviceCommandSafe('insertWordArt', {
    text: 'ONLYOFFICE',
    fontSize: 36,
    bold: true,
    caps: true,
    fontFamily: 'Arial Black',
    textColor: [255, 255, 255],     // 白色文字
    fillColor: [76, 175, 80],       // 绿色填充
    strokeColor: [46, 125, 50],     // 深绿色描边
    strokeWidth: 2,
    transform: 'textArchUp',        // 拱形向上
    width: 180,
    height: 60,
    rotation: 0
  });
}

function insertClassicWordArt() {
  serviceCommandSafe('insertPresetWordArt', {
    preset: 'classic',
    text: 'CLASSIC'
  });
}

function insertModernWordArt() {
  serviceCommandSafe('insertPresetWordArt', {
    preset: 'modern',
    text: 'MODERN'
  });
}

function insertFunWordArt() {
  serviceCommandSafe('insertPresetWordArt', {
    preset: 'fun',
    text: 'FUN!'
  });
}

// ---- Shape 测试功能 ----
function testShapeInline() {
  console.log('开始测试Shape内联插入可能性...');
  serviceCommandSafe('testShapeInline', {});
}

// ---- Shape 内联插入功能 ----
function insertRectangleInline() {
  serviceCommandSafe('insertShapeParagraph', {
    shapeType: 'rect',
    width: 80,
    height: 40,
    fillColor: [33, 150, 243],  // 蓝色
    strokeColor: [13, 71, 161],
    strokeWidth: 1,
    text: '矩形'
  });
}

function insertCircleInline() {
  serviceCommandSafe('insertShapeParagraph', {
    shapeType: 'ellipse',
    width: 60,
    height: 60,
    fillColor: [76, 175, 80],   // 绿色
    strokeColor: [46, 125, 50],
    strokeWidth: 1,
    text: '圆形'
  });
}

function insertTriangleInline() {
  serviceCommandSafe('insertShapeParagraph', {
    shapeType: 'triangle',
    width: 70,
    height: 60,
    fillColor: [255, 152, 0],   // 橙色
    strokeColor: [230, 81, 0],
    strokeWidth: 1,
    text: '三角'
  });
}

// ---- Table 表格插入功能 ----
function insertSimpleTable() {
  serviceCommandSafe('insertPresetTable', {
    preset: 'simple'
  });
}

function insertScheduleTable() {
  serviceCommandSafe('insertPresetTable', {
    preset: 'schedule'
  });
}

function insertComparisonTable() {
  serviceCommandSafe('insertPresetTable', {
    preset: 'comparison'
  });
}

function insertCustomTable() {
  serviceCommandSafe('insertTable', {
    rows: 4,
    columns: 5,
    widthType: 'percent',
    width: 100,
    headers: ['项目', 'Q1', 'Q2', 'Q3', 'Q4'],
    data: [
      ['销售额', '100万', '120万', '130万', '150万'],
      ['利润', '20万', '25万', '28万', '35万'],
      ['增长率', '10%', '20%', '8%', '15%']
    ]
  });
}

function insertDynamicTable() {
  // 模拟从系统获取动态数据
  const dynamicData = generateBusinessData();
  serviceCommandSafe('insertDynamicTable', dynamicData);
}

// 模拟业务系统数据生成
function generateBusinessData() {
  const products = ['iPhone 15', 'Samsung S24', 'Xiaomi 14', 'OPPO Find X7', 'vivo X100'];
  const months = ['1月', '2月', '3月', '4月', '5月', '6月'];

  const headers = ['产品', ...months, '总计'];
  const data = [];

  products.forEach(product => {
    const row = [product];
    let total = 0;

    // 生成随机销量数据
    for (let i = 0; i < 6; i++) {
      const sales = Math.floor(Math.random() * 1000) + 500;
      row.push(sales.toString());
      total += sales;
    }
    row.push(total.toString());
    data.push(row);
  });

  // 添加合计行
  const totalRow = ['合计'];
  for (let i = 1; i < headers.length; i++) {
    const columnTotal = data.reduce((sum, row) => sum + parseInt(row[i]), 0);
    totalRow.push(columnTotal.toString());
  }
  data.push(totalRow);

  return {
    title: '产品销量统计表',
    headers,
    data,
    metadata: {
      generatedAt: new Date().toLocaleString('zh-CN'),
      dataSource: 'ERP系统',
      category: '销售报表'
    },
    styling: {
      headerStyle: 'bold',
      totalRowStyle: 'bold',
      alternateRowColors: true
    }
  };
}

// ---- Chart Binding 图表绑定功能 - 新增 ----

// 绑定图表数据
function bindChartData() {
  const chartData = generateSampleChartData();
  serviceCommandSafe('bindChartData', chartData);
}

// 获取图表绑定摘要
function getChartSummary() {
  serviceCommandSafe('getChartSummary', {});
}

// 测试图表检测
function testChartDetection() {
  serviceCommandSafe('chartClicked', {});
}

// 生成示例图表数据
function generateSampleChartData() {
  return {
    data: {
      title: '销售趋势图',
      type: 'line-chart',
      dataSource: 'ERP系统',
      category: '销售分析',
      period: '2024年第一季度',
      metrics: {
        totalSales: 1250000,
        growthRate: 15.8,
        topProduct: 'iPhone 15',
        targetAchievement: 125
      },
      series: [
        { name: '销售额', data: [120000, 135000, 128000, 142000] },
        { name: '目标', data: [125000, 130000, 135000, 140000] }
      ],
      categories: ['1月', '2月', '3月', '4月']
    },
    metadata: {
      bindingId: 'chart_' + Date.now(),
      boundAt: new Date().toISOString(),
      dataVersion: '1.0',
      refreshInterval: 3600, // 秒
      lastUpdated: new Date().toLocaleString('zh-CN'),
      permissions: {
        canEdit: true,
        canRefresh: true,
        canExport: true
      }
    }
  };
}

// 分析选中内容
function analyzeSelection() {
  serviceCommandSafe('analyzeSelection', {});
}

// 绑定为文本字段
function bindAsTextField() {
  const fieldName = prompt('请输入字段名称:', 'customer_name') || 'text_field';
  serviceCommandSafe('bindSelection', {
    type: 'text-field',
    fieldName: fieldName,
    dataType: 'text',
    category: 'data-field'
  });
}

// 绑定为模板变量
function bindAsVariable() {
  const varName = prompt('请输入变量名称:', 'my_variable') || 'custom_var';
  serviceCommandSafe('bindSelection', {
    type: 'template-variable',
    fieldName: varName,
    category: 'template',
    metadata: {
      variable: varName
    }
  });
}

// 绑定为数据源
function bindAsDataSource() {
  const sourceName = prompt('请输入数据源名称:', 'data_table') || 'data_source';
  serviceCommandSafe('bindSelection', {
    type: 'table-data-source',
    fieldName: sourceName,
    category: 'data-binding'
  });
}

// 处理链接点击事件
function handleLinkClick(data: any) {
  console.log('处理链接点击事件，收到数据:', data)

  if (data && typeof data === 'object') {
    // 显示详细的控件信息和数据
    const info = `链接控件被点击！

      控件ID: ${data.controlId || 'unknown'}
      控件标题: ${data.controlTitle || '无标题'}
      控件Tag: ${data.tag || '无Tag'}

      绑定数据:
      ${JSON.stringify(data.data, null, 2)}

      完整信息:
      ${JSON.stringify(data, null, 2)}`;

    alert(info);

    // 你可以根据数据执行相应的操作
    // 例如：打开弹窗、跳转页面、发送请求等
    if (data.data && data.data.name) {
      console.log('检测到用户名:', data.data.name);
    }
  }
}

// 处理 WordArt 插入成功事件
function handleWordArtInserted(data: any) {
  console.log('WordArt 插入成功，数据:', data)

  if (data && data.success) {
    // 显示成功信息
    const info = `WordArt 插入成功！

      文本: ${data.parameters?.text || '未知'}
      字体大小: ${data.parameters?.fontSize || '未知'}
      字体: ${data.parameters?.fontFamily || '未知'}
      变换效果: ${data.parameters?.transform || '未知'}

${data.message || ''}`;

    // 使用更友好的通知方式
    if (confirm(info + '\n\n点击"确定"可继续插入更多WordArt')) {
      // 可以打开WordArt配置面板或其他操作
      console.log('用户选择继续操作WordArt')
    }
  }
}

// 处理 WordArt 插入错误事件
function handleWordArtError(data: any) {
  console.error('WordArt 插入失败，错误:', data)

  const errorMsg = data?.error || '未知错误'
  const info = `WordArt 插入失败！

      错误信息: ${errorMsg}

      可能的原因:
      - 当前位置不支持插入WordArt
      - API参数错误
      - OnlyOffice版本不支持WordArt功能

      请检查控制台获取详细错误信息。`;

  alert(info);
}


// 处理表格点击事件
function handleTableClick(data: any) {
  console.log('处理表格点击事件，收到数据:', data)

  if (data && data.success && data.data) {
    const tableInfo = data.data;

    // 显示详细的表格信息
    let info = `表格被点击！\n\n`;

    if (tableInfo.clickType === 'table') {
      info += `点击类型: 表格内容\n`;
      info += `表格索引: ${tableInfo.tableIndex >= 0 ? tableInfo.tableIndex : '未知'}\n`;
      info += `行索引: ${tableInfo.rowIndex >= 0 ? tableInfo.rowIndex : '未知'}\n`;
      info += `单元格索引: ${tableInfo.cellIndex >= 0 ? tableInfo.cellIndex : '未知'}\n`;
      info += `单元格文本: "${tableInfo.cellText || '空'}"\n`;
      info += `点击时间: ${tableInfo.timestamp}\n`;

      if (tableInfo.tableData) {
        info += `\n表格数据结构:\n`;
        info += `- 行数: ${tableInfo.tableData.rows}\n`;
        info += `- 列数: ${tableInfo.tableData.columns}\n`;

        if (tableInfo.tableData.content && tableInfo.tableData.content.length > 0) {
          info += `\n表格内容预览:\n`;
          tableInfo.tableData.content.slice(0, 3).forEach((row: any[], rowIndex: number) => {
            info += `第${rowIndex + 1}行: [${row.join(', ')}]\n`;
          });
          if (tableInfo.tableData.content.length > 3) {
            info += `... 还有 ${tableInfo.tableData.content.length - 3} 行\n`;
          }
        }
      }
    } else {
      info += `点击类型: 非表格区域\n`;
      info += `点击时间: ${tableInfo.timestamp}\n`;
    }

    alert(info);

    // 你可以根据表格数据执行相应的操作
    // 例如：显示表格编辑面板、发送数据到服务器等
    if (tableInfo.clickType === 'table' && tableInfo.tableData) {
      console.log('检测到表格点击，可以执行相关操作:', {
        tableIndex: tableInfo.tableIndex,
        tableData: tableInfo.tableData
      });
    }
  } else {
    console.log('表格点击检测失败:', data?.error || '未知错误');
    alert(`表格点击检测失败: ${data?.error || '未知错误'}`);
  }
}

// 处理选中内容分析结果
function handleSelectionAnalyzed(data: any) {
  console.log('处理选中内容分析结果，收到数据:', data)

  if (data && data.success && data.data) {
    const selectionInfo = data.data;

    let info = `选中内容分析结果！\n\n`;
    info += `选择类型: ${selectionInfo.selectionType}\n`;
    info += `是否可绑定: ${selectionInfo.bindable ? '是' : '否'}\n`;
    info += `选中内容: "${selectionInfo.content.substring(0, 50)}${selectionInfo.content.length > 50 ? '...' : ''}"\n\n`;

    if (selectionInfo.suggestedBindings && selectionInfo.suggestedBindings.length > 0) {
      info += `建议的绑定方式:\n`;
      selectionInfo.suggestedBindings.slice(0, 3).forEach((binding: any, index: number) => {
        info += `${index + 1}. ${binding.description} (${binding.category})\n`;
      });

      if (selectionInfo.suggestedBindings.length > 3) {
        info += `... 还有 ${selectionInfo.suggestedBindings.length - 3} 种绑定方式\n`;
      }
    }

    info += `\n分析时间: ${selectionInfo.timestamp}`;

    alert(info);

    // 可以根据分析结果执行后续操作
    if (selectionInfo.bindable) {
      console.log('内容可以绑定，建议绑定方式:', selectionInfo.suggestedBindings);
    }
  } else {
    console.log('选中内容分析失败:', data?.error || '未知错误');
    alert(`选中内容分析失败: ${data?.error || '未知错误'}`);
  }
}

// 处理选中内容绑定完成
function handleSelectionBound(data: any) {
  console.log('处理选中内容绑定完成，收到数据:', data)

  if (data && data.success) {
    const bindingInfo = data.data;

    let info = `数据绑定完成！\n\n`;
    info += `绑定方法: ${bindingInfo.method}\n`;
    info += `绑定类型: ${bindingInfo.binding?.type || '未知'}\n`;

    if (bindingInfo.binding) {
      if (bindingInfo.binding.fieldName) {
        info += `字段名称: ${bindingInfo.binding.fieldName}\n`;
      }
      if (bindingInfo.binding.dataType) {
        info += `数据类型: ${bindingInfo.binding.dataType}\n`;
      }
      if (bindingInfo.binding.originalValue) {
        info += `原始值: "${bindingInfo.binding.originalValue}"\n`;
      }
    }

    info += `\n${bindingInfo.message}`;

    alert(info);

    console.log('数据绑定成功，绑定信息:', bindingInfo.binding);
  } else {
    console.log('数据绑定失败:', data?.error || '未知错误');
    alert(`数据绑定失败: ${data?.error || '未知错误'}`);
  }
}

// 处理通用元素检测事件 - 新增
function handleElementClicked(data: any) {
  console.log('🎯 处理通用元素点击事件，收到数据:', data)

  if (data && data.success && data.data) {
    const elementInfo = data.data;

    let info = `🎯 元素点击检测结果！\n\n`;
    info += `点击类型: ${elementInfo.clickType}\n`;
    info += `检测时间: ${elementInfo.timestamp}\n\n`;

    // 根据不同元素类型显示详细信息
    if (elementInfo.clickType === 'table') {
      info += `📊 表格信息:\n`;
      info += `- 表格索引: ${elementInfo.elementInfo.tableIndex}\n`;
      info += `- 总行数: ${elementInfo.elementInfo.totalRows}\n`;
      info += `- 总列数: ${elementInfo.elementInfo.totalColumns}\n`;

      if (elementInfo.elementInfo.clickedCell) {
        info += `- 点击位置: 第${elementInfo.elementInfo.clickedCell.row}行第${elementInfo.elementInfo.clickedCell.column}列\n`;
        info += `- 单元格内容: "${elementInfo.elementInfo.clickedCellInfo?.content || '空'}"\n`;
      }

      if (elementInfo.elementInfo.sampleContent && elementInfo.elementInfo.sampleContent.length > 0) {
        info += `\n📋 表格内容预览:\n`;
        elementInfo.elementInfo.sampleContent.slice(0, 2).forEach((row: any[], index: number) => {
          info += `第${index + 1}行: [${row.slice(0, 3).join(', ')}${row.length > 3 ? '...' : ''}]\n`;
        });
      }

    } else if (elementInfo.clickType === 'paragraph') {
      info += `📝 段落信息:\n`;
      info += `- 段落索引: ${elementInfo.elementInfo.paragraphIndex}\n`;
      info += `- 对齐方式: ${elementInfo.elementInfo.alignment}\n`;
      info += `- 文本长度: ${elementInfo.elementInfo.metadata?.fullTextLength || 0}字符\n`;
      info += `- 文本预览: "${elementInfo.elementInfo.text}"\n`;

    } else if (elementInfo.clickType === 'shape') {
      info += `🔷 图形信息:\n`;
      info += `- 图形索引: ${elementInfo.elementInfo.shapeIndex}\n`;
      info += `- 图形类型: ${elementInfo.elementInfo.shapeType}\n`;
      info += `- 包含内容: ${elementInfo.elementInfo.hasContent ? '是' : '否'}\n`;

    } else if (elementInfo.clickType === 'document') {
      info += `📄 文档区域:\n`;
      info += `- 有选区: ${elementInfo.elementInfo.hasSelection ? '是' : '否'}\n`;
      info += `- 扫描元素: ${elementInfo.elementInfo.metadata?.totalElementsScanned || 0}个\n`;
    }

    // 显示扫描统计
    if (elementInfo.fullScanResults) {
      info += `\n🔍 文档扫描统计:\n`;
      info += `- 表格: ${elementInfo.fullScanResults.tablesFound}个\n`;
      info += `- 段落: ${elementInfo.fullScanResults.paragraphsFound}个\n`;
      info += `- 图形: ${elementInfo.fullScanResults.shapesFound}个\n`;
    }

    alert(info);

    // 可以根据元素类型执行不同的操作
    console.log('🎯 元素类型:', elementInfo.clickType, '详细信息:', elementInfo.elementInfo);

  } else {
    console.log('❌ 通用元素检测失败:', data?.error || '未知错误');
    // 不显示失败的alert，避免过多弹窗
  }
}

// 处理精确表格单元格检测事件 - 新增
function handlePreciseTableCellClicked(data: any) {
  console.log('📊 处理精确表格单元格点击事件，收到数据:', data)

  if (data && data.success && data.data) {
    const cellInfo = data.data;

    let info = `📊 精确表格单元格点击！\n\n`;
    info += `🎯 单元格位置:\n`;
    info += `- 行: 第${cellInfo.cellPosition.row}行\n`;
    info += `- 列: 第${cellInfo.cellPosition.column}列\n`;
    info += `- 单元格内容: "${cellInfo.cellContent}"\n\n`;

    info += `📋 表格信息:\n`;
    info += `- 总行数: ${cellInfo.tableInfo.totalRows}\n`;
    info += `- 总列数: ${cellInfo.tableInfo.totalColumns}\n`;
    info += `- 检测方法: ${cellInfo.detectionMethod}\n`;
    info += `- 检测时间: ${cellInfo.timestamp}\n`;

    alert(info);

    // 可以根据精确位置执行特定操作
    console.log('📊 精确单元格信息:', {
      row: cellInfo.cellPosition.row,
      column: cellInfo.cellPosition.column,
      content: cellInfo.cellContent,
      tableInfo: cellInfo.tableInfo
    });

    // 示例：根据点击的单元格执行不同操作
    if (cellInfo.cellPosition.row === 1) {
      console.log('🎯 点击了表头行，可以执行排序等操作');
    } else {
      console.log('🎯 点击了数据行，可以执行编辑等操作');
    }

  } else {
    console.log('❌ 精确表格单元格检测失败:', data?.error || '未知错误');
    // 这通常是因为不是表格区域或检测方法失败，不需要alert
  }
}
// 处理图表点击事件 - 新增
function handleChartClicked(data: any) {
  console.log('📈 处理图表点击事件，收到数据:', data)

  if (data && data.success && data.data) {
    const chartInfo = data.data;

    let info = `📈 图表点击检测结果！\n\n`;
    info += `点击类型: ${chartInfo.clickType}\n`;
    info += `检测时间: ${chartInfo.timestamp}\n\n`;

    if (chartInfo.chartInfo) {
      info += `📊 图表信息:\n`;
      info += `- 图表索引: ${chartInfo.chartInfo.chartIndex}\n`;
      info += `- 图表类型: ${chartInfo.chartInfo.chartType}\n`;

      // 显示详细图表类型信息
      if (chartInfo.chartInfo.detailedChartType) {
        const detailed = chartInfo.chartInfo.detailedChartType;
        info += `- 图表分类: ${detailed.category}\n`;
        info += `- 具体类型: ${detailed.specificType}\n`;
        info += `- 类型描述: ${detailed.description}\n`;
        info += `- 识别置信度: ${(detailed.confidence * 100).toFixed(1)}%\n`;

        if (detailed.properties && Object.keys(detailed.properties).length > 0) {
          info += `- 属性信息: ${JSON.stringify(detailed.properties)}\n`;
        }
      }

      info += `- 有绑定数据: ${chartInfo.chartInfo.hasBindingData ? '是' : '否'}\n`;
      info += `- 绑定方法: ${chartInfo.chartInfo.bindingMethod || '无'}\n\n`;
    }

    if (chartInfo.boundData) {
      info += `💾 绑定的数据:\n`;
      if (chartInfo.boundData.title) {
        info += `- 标题: ${chartInfo.boundData.title}\n`;
      }
      if (chartInfo.boundData.type) {
        info += `- 类型: ${chartInfo.boundData.type}\n`;
      }
      if (chartInfo.boundData.dataSource) {
        info += `- 数据源: ${chartInfo.boundData.dataSource}\n`;
      }
      if (chartInfo.boundData.period) {
        info += `- 时间周期: ${chartInfo.boundData.period}\n`;
      }

      if (chartInfo.boundData.metrics) {
        info += `\n📈 关键指标:\n`;
        const metrics = chartInfo.boundData.metrics;
        if (metrics.totalSales) {
          info += `- 总销售额: ¥${metrics.totalSales.toLocaleString()}\n`;
        }
        if (metrics.growthRate) {
          info += `- 增长率: ${metrics.growthRate}%\n`;
        }
        if (metrics.topProduct) {
          info += `- 热门产品: ${metrics.topProduct}\n`;
        }
        if (metrics.targetAchievement) {
          info += `- 目标达成: ${metrics.targetAchievement}%\n`;
        }
      }

      if (chartInfo.boundData.series && chartInfo.boundData.series.length > 0) {
        info += `\n📊 数据系列:\n`;
        chartInfo.boundData.series.slice(0, 2).forEach((series: any) => {
          info += `- ${series.name}: [${series.data.slice(0, 3).join(', ')}${series.data.length > 3 ? '...' : ''}]\n`;
        });
      }
    }

    if (chartInfo.bindingMetadata) {
      info += `\n🔗 绑定信息:\n`;
      info += `- 绑定ID: ${chartInfo.bindingMetadata.bindingId}\n`;
      info += `- 绑定时间: ${chartInfo.bindingMetadata.boundAt}\n`;
      info += `- 绑定方法: ${chartInfo.bindingMetadata.bindingMethod}\n`;
    }

    if (chartInfo.detectionSummary) {
      info += `\n🔍 检测摘要:\n`;
      info += `- 文档中图表总数: ${chartInfo.detectionSummary.totalChartsFound}\n`;
      info += `- 有数据的图表: ${chartInfo.detectionSummary.chartsWithData}\n`;
      info += `- 有选区: ${chartInfo.detectionSummary.hasSelection ? '是' : '否'}\n`;
    }

    alert(info);

    // 可以根据图表数据执行相应的操作
    console.log('📈 图表绑定数据:', chartInfo.boundData);

    // 示例：根据图表类型执行不同操作
    if (chartInfo.boundData) {
      if (chartInfo.boundData.type === 'line-chart') {
        console.log('📈 这是线图，可以显示趋势分析');
      } else if (chartInfo.boundData.type === 'bar-chart') {
        console.log('📊 这是柱状图，可以显示比较分析');
      }

      // 可以发送数据到外部系统进行处理
      if (chartInfo.boundData.dataSource === 'ERP系统') {
        console.log('🔄 可以触发ERP系统数据刷新');
      }
    }

  } else if (data && data.success === false) {
    console.log('❌ 图表检测失败或无图表:', data.error || '未知错误');
    // 不显示错误alert，避免过多弹窗
  }
}

function handleBindingClicked(data: any) {
  console.log('处理绑定控件点击事件，收到数据:', data)

  if (data && data.success && data.data) {
    const bindingInfo = data.data

    let info = `绑定控件被点击！\\n\\n`
    info += `点击类型: ${bindingInfo.clickType}\\n`
    info += `绑定类型: ${bindingInfo.bindingType}\\n`
    info += `控件索引: ${bindingInfo.controlIndex}\\n`
    info += `控件别名: ${bindingInfo.controlAlias || '无'}\\n`
    info += `控件内容: "${bindingInfo.controlContent || '空'}"\\n`
    info += `点击时间: ${bindingInfo.timestamp}\\n`

    if (bindingInfo.bindingData) {
      info += `\\n绑定数据详情:\\n`

      if (bindingInfo.bindingType === 'data-binding') {
        info += `- 字段名称: ${bindingInfo.bindingData.fieldName}\\n`
        info += `- 数据类型: ${bindingInfo.bindingData.dataType}\\n`
        info += `- 原始值: "${bindingInfo.bindingData.originalValue}"\\n`
        info += `- 绑定时间: ${bindingInfo.bindingData.boundAt}\\n`
      } else if (bindingInfo.bindingType === 'template-variable') {
        info += `- 变量名: ${bindingInfo.bindingData.variableName}\\n`
        info += `- 原始值: "${bindingInfo.bindingData.originalValue}"\\n`
        info += `- 绑定时间: ${bindingInfo.bindingData.boundAt}\\n`
      } else if (bindingInfo.bindingType === 'table-data-binding') {
        info += `- 表格名称: ${bindingInfo.bindingData.tableName}\\n`
        info += `- 绑定模式: ${bindingInfo.bindingData.bindingMode}\\n`
        info += `- 绑定时间: ${bindingInfo.bindingData.boundAt}\\n`
      } else if (bindingInfo.bindingType === 'paragraph-template') {
        info += `- 模板名称: ${bindingInfo.bindingData.templateName}\\n`
        info += `- 原始内容: "${bindingInfo.bindingData.originalContent?.substring(0, 50)}${bindingInfo.bindingData.originalContent?.length > 50 ? '...' : ''}"\\n`
        info += `- 绑定时间: ${bindingInfo.bindingData.boundAt}\\n`
      } else if (bindingInfo.bindingType === 'custom-binding') {
        info += `- 自定义类型: ${bindingInfo.bindingData.customType}\\n`
        info += `- 字段名称: ${bindingInfo.bindingData.fieldName}\\n`
        info += `- 原始值: "${bindingInfo.bindingData.originalValue}"\\n`
        info += `- 绑定时间: ${bindingInfo.bindingData.boundAt}\\n`
      }
    }

    alert(info)

    // 可以根据绑定数据执行相应的操作
    console.log('检测到绑定控件点击，可以执行相关操作:', {
      bindingType: bindingInfo.bindingType,
      bindingData: bindingInfo.bindingData
    })

    // 根据不同的绑定类型执行不同的操作
    if (bindingInfo.bindingType === 'data-binding') {
      // 可以打开数据编辑面板
      console.log('数据字段绑定被点击，可以打开数据编辑界面')
    } else if (bindingInfo.bindingType === 'template-variable') {
      // 可以打开变量设置面板
      console.log('模板变量绑定被点击，可以打开变量设置界面')
    } else if (bindingInfo.bindingType === 'table-data-binding') {
      // 可以打开数据源配置面板
      console.log('表格数据源绑定被点击，可以打开数据源配置界面')
    }
  } else {
    console.log('绑定控件点击检测失败或无绑定控件:', data?.error || '未知错误')
    if (data?.data?.clickType === 'non-binding') {
      console.log('当前点击位置不是绑定控件')
    }
  }
}


onMounted(() => createEditor())
onBeforeUnmount(() => {
  if (editor?.destroyEditor) editor.destroyEditor()
  if (flushTimer !== null) clearInterval(flushTimer)
})
</script>
