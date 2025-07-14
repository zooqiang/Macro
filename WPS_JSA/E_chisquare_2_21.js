Attribute Module_Name = "E_chisquare_2_21"
function E_chiSquare_2x2_TestLocal() {
	// 如果验证失败（用户取消或未授权），函数会直接返回
      if (!checkZouAndVerifyUser()) {
        alert("操作终止：未通过验证");
        return;
    }    
   // 只有验证通过或已有"邹"字样才会继续执行 
    var selection = Application.Selection;
    if (!selection || !selection.Text) {
        alert("请先选中要计算卡方检验的单元格。");
        return;
    }

    var cellText1, cellText2, cellText3, cellText4;
    var A, B, C, D, O, p, x, Y;
    var chiSquare, pValue, result;

    // 检查选中的单元格数量
    var cellCount = selection.Cells.Count;
    if (cellCount !== 2 && cellCount !== 4) {
        alert("请选中两个或四个单元格！");
        return;
    }

    // 模式一：直接提取 A, B, C, D
    if (cellCount === 4) {
        cellText1 = selection.Cells.Item(1).Range.Text;
        cellText2 = selection.Cells.Item(2).Range.Text;
        cellText3 = selection.Cells.Item(3).Range.Text;
        cellText4 = selection.Cells.Item(4).Range.Text;

        // 去除表格符号（如段落标记）并转换为半角
        cellText1 = cellText1.replace(/\r/g, "").normalize("NFKC");
        cellText2 = cellText2.replace(/\r/g, "").normalize("NFKC");
        cellText3 = cellText3.replace(/\r/g, "").normalize("NFKC");
        cellText4 = cellText4.replace(/\r/g, "").normalize("NFKC");

        // 提取数字（A, B, C, D）
        A = extractNumber(cellText1);
        B = extractNumber(cellText2);
        C = extractNumber(cellText3);
        D = extractNumber(cellText4);
    }
    // 模式二：提取 A, B, O, P，计算 X, Y, C, D
    else if (cellCount === 2) {
        cellText1 = selection.Cells.Item(1).Range.Text;
        cellText2 = selection.Cells.Item(2).Range.Text;

        // 去除表格符号（如段落标记）并转换为半角
        cellText1 = cellText1.replace(/\r/g, "").normalize("NFKC");
        cellText2 = cellText2.replace(/\r/g, "").normalize("NFKC");

        // 提取括号前的数字（A 和 B）
        A = extractNumberBeforeParenthesis(cellText1);
        B = extractNumberBeforeParenthesis(cellText2);

        // 提取括号后的数字（O 和 P）
        O = extractNumberAfterParenthesis(cellText1);
        p = extractNumberAfterParenthesis(cellText2);

        // 计算 X, Y, C, D
        x = Math.round((A / O) * 100);
        Y = Math.round((B / p) * 100);
        C = x - A;
        D = Y - B;
    }

    // 检查是否都为正整数
    if ([A, B, C, D].some(isNaN)) {
        alert("数据无效，请确保单元格中包含有效的数字。");
        return;
    }

    // 计算卡方值
    var N = A + B + C + D;
    var numerator = Math.pow(A * D - B * C, 2) * N;
    var denominator = (A + B) * (C + D) * (A + C) * (B + D);

    if (denominator === 0) {
        alert("分母为零，无法计算卡方值。");
        return;
    }

    chiSquare = numerator / denominator;

    // 本地计算 p 值（自由度=1）
    pValue = chi2cdf(chiSquare, 1);

    // 构建结果字符串
    result = "A: " + A + "\r\n" +
             "B: " + B + "\r\n" +
             "C: " + C + "\r\n" +
             "D: " + D + "\r\n" +
             "卡方值: " + chiSquare.toFixed(4) + "\r\n" +
             "P 值: " + pValue.toFixed(4);

    // 添加批注
    selection.Comments.Add(selection.Range, result + "\n");
}

// 自由度为1的卡方分布 p 值计算（近似）
function chi2cdf(x, df) {
    // 近似公式：适用于 df = 1
    return 1 - erf(Math.sqrt(x / 2));
}

// 实现 erf 函数（误差函数）
function erf(x) {
    // 用近似多项式展开法实现 erf
    const a1 = 0.254829592;
    const a2 = -0.284496736;
    const a3 = 1.421413741;
    const a4 = -1.453152027;
    const a5 = 1.061405429;
    const p = 0.3275911;

    var sign = (x < 0) ? -1 : 1;
    x = Math.abs(x);

    var t = 1.0 / (1.0 + p * x);
    var y = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.exp(-x * x);

    return sign * y;
}

// 提取纯数字
function extractNumber(text) {
    var numStr = text.replace(/[^0-9.]/g, "");
    return numStr ? parseFloat(numStr) : 0;
}

// 提取括号前的数字
function extractNumberBeforeParenthesis(text) {
    var pos = text.indexOf("(");
    if (pos > 0) {
        return parseFloat(text.substring(0, pos).trim());
    }
    return 0;
}

// 提取括号内的数字
function extractNumberAfterParenthesis(text) {
    var posStart = text.indexOf("(");
    var posEnd = text.indexOf(")");
    if (posStart > 0 && posEnd > posStart) {
        var numStr = text.substring(posStart + 1, posEnd).trim().replace("%", "");
        return parseFloat(numStr);
    }
    return 0;
}