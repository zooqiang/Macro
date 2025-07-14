Attribute Module_Name = "F_comments_del"
function DeleteAllComments() {
    // 获取当前活动文档
    var doc = Application.ActiveDocument;
    
    // 检查文档中是否有批注
    if (doc.Comments.Count === 0) {
        Application.Alert("文档中没有批注可删除。");
        return;
    }
    
    // 使用WPS提供的对话框进行确认
    var result = MsgBox(
        "确定要删除文档中的所有批注吗？共" + doc.Comments.Count + "个批注。", 
        4, // 4表示"是/否"按钮
        "删除批注确认"
    );
    
    // 6表示点击了"是"，其他值表示"否"或关闭对话框
    if (result !== 6) {
        return;
    }
    
    // 记录原始批注数量
    var originalCount = doc.Comments.Count;
    
    // 循环删除所有批注（从后往前删除更安全）
    while (doc.Comments.Count > 0) {
        doc.Comments.Item(doc.Comments.Count).Delete();
    }
    
    // 显示完成消息
    alert("已删除所有批注，共删除" + originalCount + "个批注。");
}