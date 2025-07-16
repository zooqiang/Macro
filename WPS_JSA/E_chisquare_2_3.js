Attribute Module_Name = "E_chisquare_2_3"
function E_chiSquareTest_2x3() {
    var selection = Application.Selection;
    if (!selection) {
        alert("请先选中表格中的单元格。");
        return;
    }

    var cellCount = selection.Cells.Count;

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
    } else {
        alert("请选择 3 个或 6 个单元格！");
        return;
    }

    // 检查是否都为有效数字
    if (observed.some(isNaN)) {
        alert("数据无效，请确保单元格中包含有效的数字。");
        return;
    }

    // 构建 2x3 矩阵
    var table = [];
    if (cellCount === 6) {
        table = [
            [observed[0], observed[1], observed[2]],
            [observed[3], observed[4], observed[5]]
        ];
    } else if (cellCount === 3) {
        table = [
            [observed[0], observed[2], observed[4]], // A1, A2, A3
            [observed[1], observed[3], observed[5]]  // B1, B2, B3
        ];
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

// 卡方分布累积函数（近似）
function chi2cdf(x, df) {
    if (df === 1) return 1 - erf(Math.sqrt(x / 2));
    else if (df === 2) return 1 - Math.exp(-x / 2);
    else if (df === 3) return gamma_p(1.5, x / 2);
    else return chi2cdfApprox(x, df);
}

function chi2cdfApprox(x, df) {
    var z = Math.sqrt(2 * df);
    var t = (Math.pow(x / df, 1 / 3) - (1 - 2 / (9 * df))) / (Math.sqrt(2) / (3 * Math.sqrt(2 / (9 * df))));
    return 1 - erf(Math.abs(t) / Math.sqrt(2));
}

// 误差函数 erf(x)
function erf(x) {
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

// 伽马函数 P 分布（用于 df=3）
function gamma_p(a, x) {
    var gln = lngamma(a);
    if (x <= 0) return 0;
    else if (x < a + 1) return gammap_series(a, x, gln);
    else return 1 - gammac_continued(a, x, gln);
}

function lngamma(x) {
    var cof = [76.18009173, -86.50532033, 24.01409822, -1.231739516, 0.120858003e-2, -0.536382e-5];
    var y = x;
    var tmp = x + 5.5;
    tmp -= (x + 0.5) * Math.log(tmp);
    var ser = 1.000000000190015;
    for (var j = 0; j < 6; j++) ser += cof[j] / (++y);
    return -tmp + Math.log(2.5066282746310005 * ser / x);
}

function gammap_series(a, x, gln) {
    var eps = 1e-7;
    var itmax = 100;
    var n;
    var sum = 1 / a;
    var add = sum;
    for (n = 1; n <= itmax; n++) {
        add *= x / (a + n);
        sum += add;
        if (Math.abs(add) < Math.abs(sum) * eps) break;
    }
    return Math.exp(-x + a * Math.log(x) - gln) * sum;
}

function gammac_continued(a, x, gln) {
    var eps = 1e-7;
    var itmax = 100;
    var FPMIN = 1e-30;
    var b = x + 1 - a;
    var c = 1 / FPMIN;
    var d = 1 / b;
    var h = d;
    for (var i = 1; i <= itmax; i++) {
        var an = -i * (i - a);
        b += 2;
        d = an * d + b;
        if (Math.abs(d) < FPMIN) d = FPMIN;
        c = b + an / c;
        if (Math.abs(c) < FPMIN) c = FPMIN;
        d = 1 / d;
        var del = d * c;
        h *= del;
        if (Math.abs(del - 1) < eps) break;
    }
    return Math.exp(-x + a * Math.log(x) - gln) * h;
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

// 提取括号内的数字（百分比对应的总数）
function extractNumberAfterParenthesis(text) {
    var posStart = text.indexOf("(");
    var posEnd = text.indexOf(")");
    if (posStart > 0 && posEnd > posStart) {
        var numStr = text.substring(posStart + 1, posEnd).trim().replace("%", "");
        return Math.round(parseFloat(numStr) * 10); // 如 50% → 10 → 总数=20
    }
    return 0;
}