/* globals window, Asc, Api */

// 生命周期：初始化（编辑器加载完后会回调）
window.Asc.plugin.init = function () {
  // 你可以在这里做准备动作（无 UI 插件也可只留空）
  console.log('[InsertHello] init')
  try { window.Asc.plugin.executeMethod('ShowMessage', ['InsertHello 插件已就绪']); } catch (e) {}
};

// 生命周期：按钮点击（在 config.json 的 buttons 数组里定义的按钮）
window.Asc.plugin.button = function (id) {
  console.log('[InsertHello] button clicked, id=', id);

  // 在文档上下文中执行：插入段落 + 文本
  window.Asc.plugin.callCommand(function () {
    var doc = Api.GetDocument();
    var p = Api.CreateParagraph();
    p.AddText('{{name}}');
    doc.InsertContent([p]);
  }, true);

  // 操作完成后关闭插件（可选）
  // window.Asc.plugin.executeMethod('Close', []);
};

// 插件关闭回调（生命周期：结束/销毁前）
window.Asc.plugin.onPluginClosed = function () {
  console.log('[InsertHello] closed');
};
