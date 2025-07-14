

    var selection = Application.Selection;
    if (!selection) {
        alert("请先选中表格中的单元格。");
        return;
    }

    var cellCount = selection.Cells.Count;
    var rowCount = selection.Rows.Count;
    var colCount = selection.Columns.Count;

    // 存放观测值
    var observed = [];

    // 模式一：选择 6 个单元格，直接提取数字
    if (cellCount === 6) {
        for (var i = 1; i <= 6; i++) {
            var cellText = selection.Cells.Item(i).Range.Text;
            cellText = cellText.replace(/\r/g, "").normalize("NFKC"); // 清理文本
            var value = extractNumber(cellText);
            observed.push(value);
        }
        
        // 根据行列形状构建2x3矩阵
        if (rowCount === 2 && colCount === 3) {
            // 2行3列：按行读取
            table = [
                [observed[0], observed[1], observed[2]],
                [observed[3], observed[4], observed[5]]
            ];
        } else if (rowCount === 3 && colCount === 2) {
            // 3行2列：按列读取（转置）
            table = [
                [observed[0], observed[2], observed[4]], // 第一列数据
                [observed[1], observed[3], observed[5]]  // 第二列数据
            ];
        } else if (rowCount === 6 && colCount === 1) {
            // 6行1列：前3行为第一组，后3行为第二组
            table = [
                [observed[0], observed[1], observed[2]],
                [observed[3], observed[4], observed[5]]
            ];
        } else if (rowCount === 1 && colCount === 6) {
            // 1行6列：前3列为第一组，后3列为第二组
            table = [
                [observed[0], observed[1], observed[2]],
                [observed[3], observed[4], observed[5]]
            ];
        } else {
            alert("请选择2×3、3×2、6×1或1×6的单元格区域！");
            return;
        }
    }
    // 模式二：选择 3 个单元格，格式为 "n (N%)"
    else if (cellCount === 3) {
        for (var i = 1; i <= 3; i++) {
            var cellText = selection.Cells.Item(i).Range.Text;
            cellText = cellText.replace(/\r/g, "").normalize("NFKC");

            var success = extractNumberBeforeParenthesis(cellText); // 提取括号前
            var total = extractNumberAfterParenthesis(cellText);     // 提取括号内百分比对应总数

            var fail = total - success;

            // 添加两组数据：成功 & 失败
            observed.push(success);
            observed.push(fail);
        }
        // 固定构建2x3矩阵（3组数据转置）
        table = [
            [observed[0], observed[2], observed[4]], // A1, A2, A3
            [observed[1], observed[3], observed[5]]  // B1, B2, B3
        ];
    } else {
        alert("请选择 3 个或 6 个单元格！");
        return;
    }

    // 检查是否都为有效数字
    if (observed.some(isNaN)) {
        alert("数据无效，请确保单元格中包含有效的数字。");
        return;
    }

    // 计算行总和、列总和、总样本量
    var rowTotal = [
        table[0][0] + table[0][1] + table[0][2],
        table[1][0] + table[1][1] + table[1][2]
    ];
    var colTotal = [
        table[0][0] + table[1][0],
        table[0][1] + table[1][1],
        table[0][2] + table[1][2]
    ];
    var total = rowTotal[0] + rowTotal[1];

    // 计算期望值 & 卡方值
    var chiSquare = 0;
    for (var i = 0; i < 2; i++) {
        for (var j = 0; j < 3; j++) {
            var expected = (rowTotal[i] * colTotal[j]) / total;
            var diff = table[i][j] - expected;
            chiSquare += (diff * diff) / expected;
        }
    }

    // 自由度 df = (2-1)*(3-1) = 2
    var pValue = chi2cdf(chiSquare, 2);

    // 构造输出结果
    var result = "观测数据：\r\n";
    result += "A1: " + table[0][0] + ", A2: " + table[0][1] + ", A3: " + table[0][2] + "\r\n";
    result += "B1: " + table[1][0] + ", B2: " + table[1][1] + ", B3: " + table[1][2] + "\r\n";
    result += "卡方值: " + chiSquare.toFixed(4) + "\r\n";
    result += "P 值: " + pValue.toFixed(4) + "\r\n";
    result += "自由度: 2";

    // 添加批注
    selection.Comments.Add(selection.Range, result + "\n");
}

// 其余辅助函数保持不变（chi2cdf, erf, gamma_p 等）