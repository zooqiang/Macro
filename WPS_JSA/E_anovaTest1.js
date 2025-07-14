
function E_anovaTest() {

    try {
        // 1. 获取选中单元格
        var selection = Application.Selection;
        if (!selection || selection.Cells.Count !== 3) {
            MsgBox("请选择三个相邻的单元格（每组x±s格式）", jsOKOnly, "提示");
            return;
        }

        // 2. 提取三个单元格的文本和数值
        var groups = [];
        for (var i = 1; i <= 3; i++) {
            var cell = selection.Cells.Item(i);
            var cellText = cell.Range.Text;
            var nums = cellText.match(/[-+]?\d*\.?\d+/g) || [0, 0];
            
            groups.push({
                mean: parseFloat(nums[0]) || 0,
                sd: parseFloat(nums[1]) || 0,
                cell: cell
            });
        }

        // 3. 获取每组样本量
        function getSampleSize(prompt, defaultValue) {
            var input = InputBox(prompt, "输入样本量", defaultValue, 200, 200);
            if (input === null || input === "") throw new Error("输入取消");
            var n = parseInt(input);
            if (isNaN(n) || n <= 1) throw new Error("样本量必须>1");
            return n;
        }

        var n1 = getSampleSize("第一组样本量（n1）：", "30");
        var n2 = getSampleSize("第二组样本量（n2）：", "30");
        var n3 = getSampleSize("第三组样本量（n3）：", "30");
        
        groups[0].n = n1;
        groups[1].n = n2;
        groups[2].n = n3;

        // 4. 方差分析计算
        // 计算总平均值
        var totalN = n1 + n2 + n3;
        var grandMean = (groups[0].mean * n1 + groups[1].mean * n2 + groups[2].mean * n3) / totalN;
        
        // 计算组间平方和（SSB）
        var ssb = 0;
        groups.forEach(function(group) {
            ssb += group.n * Math.pow(group.mean - grandMean, 2);
        });
        
        // 计算组内平方和（SSW）
        var ssw = 0;
        groups.forEach(function(group) {
            ssw += (group.n - 1) * Math.pow(group.sd, 2);
        });
        
        // 计算总平方和（SST）
        var sst = ssb + ssw;
        
        // 计算自由度
        var dfBetween = 2; // k-1 (k=3组)
        var dfWithin = totalN - 3; // N-k
        var dfTotal = totalN - 1; // N-1
        
        // 计算均方
        var msBetween = ssb / dfBetween;
        var msWithin = ssw / dfWithin;
        
        // 计算F值
        var fValue = msBetween / msWithin;
        
        // 计算p值
        var pValue = fDistRightTail(fValue, dfBetween, dfWithin);

        // 5. 生成批注内容
        var commentText = 
            "【方差分析结果】\r\n" +
            "第一组（x1±s1）：" + groups[0].mean.toFixed(2) + " ± " + groups[0].sd.toFixed(2) + " (n1=" + n1 + ")\r\n" +
            "第二组（x2±s2）：" + groups[1].mean.toFixed(2) + " ± " + groups[1].sd.toFixed(2) + " (n2=" + n2 + ")\r\n" +
            "第三组（x3±s3）：" + groups[2].mean.toFixed(2) + " ± " + groups[2].sd.toFixed(2) + " (n3=" + n3 + ")\r\n" +
            "--------------------------------\r\n" +
            "组间平方和(SSB) = " + ssb.toFixed(4) + "\r\n" +
            "组内平方和(SSW) = " + ssw.toFixed(4) + "\r\n" +
            "总平方和(SST) = " + sst.toFixed(4) + "\r\n" +
            "--------------------------------\r\n" +
            "组间自由度 = " + dfBetween + "\r\n" +
            "组内自由度 = " + dfWithin + "\r\n" +
            "F值 = " + fValue.toFixed(4) + "\r\n" +
            "p值 = " + pValue.toFixed(4) + "\r\n" +
            "--------------------------------\r\n" +
            "【提示】这是单因素方差分析结果，假设各组数据独立、正态分布且方差齐性，仅供参考";

        // 添加批注到第一个单元格
       // selection.Cells.Item(1).Comments.Add(selection.Cells.Item(1).Range, commentText);
      selection.Comments.Add(selection.Range, commentText + "\n");
    } catch (e) {
        if (e.message !== "输入取消") {
            MsgBox("错误：" + e.message, jsCritical, "错误");
        }
    }
}

// F分布右尾概率计算（近似）
function fDistRightTail(f, df1, df2) {
    // 更精确的实现
    var x = df2 / (df2 + df1 * f);
    return incompleteBeta(x, df2/2, df1/2);
}

function incompleteBeta(x, a, b) {
    // 使用更稳定的算法，如Boost库中的实现
    // 这里展示概念性代码，实际应使用专业统计库
    if (x <= 0) return 0;
    if (x >= 1) return 1;
    
    // 使用对数变换避免数值问题
    var lbeta = logGamma(a) + logGamma(b) - logGamma(a + b);
    var logx = Math.log(x);
    var log1mx = Math.log(1 - x);
    
    // 系列展开计算
    var term = Math.exp(a * logx + b * log1mx - lbeta) / a;
    var sum = term;
    
    for (var k = 0; k < 1000; k++) {
        term *= x * (a + b + k) / (a + k + 1);
        sum += term;
        if (Math.abs(term) < 1e-10) break;
    }
    
    return sum;
}

// 正则化不完全Beta函数近似（简化版）
function regularizedIncompleteBeta(x, a, b) {
    // 这是非常简化的近似，实际应用中应使用更精确的算法
    // 或调用统计库中的函数
    
    // 使用连分数展开的简化版本
    var epsilon = 1e-8;
    var result = 0;
    var m = 0;
    var maxIter = 1000;
    
    while (m < maxIter) {
        var numerator;
        if (m === 0) {
            numerator = 1;
        } else if (m % 2 === 0) {
            numerator = (m/2) * (b - m/2) * x / ((a + m - 1) * (a + m));
        } else {
            numerator = -((a + (m-1)/2) * (a + b + (m-1)/2) * x) / ((a + m - 1) * (a + m));
        }
        
        var denominator = 1 + numerator / (result === 0 ? 1 : result);
        
        if (Math.abs(denominator - result) < epsilon) {
            break;
        }
        
        result = denominator;
        m++;
    }
    
    // 前置乘数
    var prefix = Math.pow(x, a) * Math.pow(1 - x, b) / (a * betaFunc(a, b));
    
    return prefix * result;
}

// Beta函数计算
function betaFunc(a, b) {
    return Math.exp(logGamma(a) + logGamma(b) - logGamma(a + b));
}

// 对数Gamma函数近似
function logGamma(x) {
    // Lanczos近似
    var cof = [
        76.18009172947146, -86.50532032941677, 24.01409824083091,
        -1.231739572450155, 0.1208650973866179e-2, -0.5395239384953e-5
    ];
    
    var ser = 1.000000000190015;
    var tmp = x + 5.5;
    tmp -= (x + 0.5) * Math.log(tmp);
    
    var y = x;
    for (var j = 0; j < 6; j++) {
        y += 1;
        ser += cof[j] / y;
    }
    
    return -tmp + Math.log(2.5066282746310005 * ser / x);
}