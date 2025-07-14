Attribute Module_Name = "F_comments_rename"
function ChangeCommentAuthorWithInput() {
    var doc = Application.ActiveDocument;
    
    // 弹出输入框获取新用户名
    var newAuthorName = InputBox("请输入新的批注用户名", "修改批注用户名", "", 1, 1);
    
    if (newAuthorName == "") {
        Application.Alert("未输入用户名，操作已取消");
        return;
    }
    
    // 遍历并修改所有批注
    for (var i = 1; i <= doc.Comments.Count; i++) {
        doc.Comments.Item(i).Author = newAuthorName;
    }
    
    alert("已将所有批注用户名修改为: " + newAuthorName);
}