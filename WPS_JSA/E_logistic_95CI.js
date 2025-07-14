
function E_logistic_95CI() {
    var selection = Application.Selection;
    
    if (!selection || !selection.Cells || selection.Cells.Count === 0) {
        MsgBox("请选择至少一个单元格", jsOKOnly, "提示");
        return;
    }
    
    var rangeText = "";
    var calculationResult = null;
    
    if (selection.Cells.Count === 2) {
        var cell1Text = getPureNumberText(selection.Cells.Item(1));
        var cell2Text = getPureNumberText(selection.Cells.Item(2));
        
        if (cell1Text === null || cell2Text === null) {
            MsgBox("单元格必须包含有效数字", jsOKOnly, "错误");
            return;
        }
        
        var lower = parseFloat(cell1Text);  // 使用统一命名：lower
        var upper = parseFloat(cell2Text);  // 使用统一命名：upper
        
        // 检查95%CI范围是否合理
        if (lower > upper) {
            MsgBox("95%CI范围错误：下限值大于上限值\n请检查输入的置信区间数据", jsOKOnly, "错误");
            return;
        }
        
        rangeText = lower + " 到 " + upper;
        calculationResult = calculateLogisticResults(lower, upper);
    } 
    else if (selection.Cells.Count === 1) {
        var text = selection.Text.trim();
        rangeText = extractSingleCellRange(text);
        
        if (!rangeText) {
            // 增强错误提示，显示实际内容
            MsgBox("未识别到有效的数值范围格式，请先修改数值区间符号\n实际内容: '" + text + "'\n支持格式：0.8-1.6、0.8~1.6、(0.8-1.6)、1.2(0.8-1.6)", jsOKOnly, "错误");
            return;
        }
        
        var rangeValues = rangeText.split(" 到 ");
        if (rangeValues.length === 2) {
            var lower = parseFloat(rangeValues[0]);
            var upper = parseFloat(rangeValues[1]);
            
            // 检查95%CI范围是否合理
            if (lower > upper) {
                MsgBox("95%CI范围错误：下限值大于上限值\n请检查输入的置信区间数据", jsOKOnly, "错误");
                return;
            }
            
            calculationResult = calculateLogisticResults(lower, upper);
        }
    }
    
    // 构建批注内容
    var commentContent = rangeText;
    if (calculationResult) {
        if (calculationResult.error) {
            commentContent += "\r\n【错误】：" + calculationResult.error;
        } else {
            commentContent += "\r\n【Logistic回归结果：】";
            commentContent += "\r\nβ: " + calculationResult.beta;
            commentContent += "\r\nS.E.: " + calculationResult.SE;
            commentContent += "\r\nWald χ²: " + calculationResult.waldChi2; // 新增 Wald χ² 值
            commentContent += "\r\nP值: " + calculationResult.pValue;
            commentContent += "\r\nOR: " + calculationResult.OR;
        }
    }
    
    // 添加批注（不显示弹窗确认）
    selection.Comments.Add(selection.Range, commentContent + "\n");
}

// 增强的计算函数（包含OR、P值和Wald χ²）
function calculateLogisticResults(lower, upper) {
    // 安全检查 - OR的95%CI不能≤0
    if (lower <= 0 || upper <= 0) {
        return { error: "OR的95%CI不能≤0" };
    }
    
    // 安全计算对数
    const safeLog = (x) => x <= 0 ? Math.log(0.0001) : Math.log(x);
    
    // 计算beta和SE
    const logLower = safeLog(lower);
    const logUpper = safeLog(upper);
    const beta = (logLower + logUpper) / 2;
    const SE = (logUpper - logLower) / 3.92; // 3.92 = 2*1.96
    
    // 计算OR值
    const OR = Math.exp(beta);
    
    // 计算 Wald χ² 值（新增）
    const waldChi2 = Math.pow(beta / SE, 2);
    
    // 计算z值
    const zScore = Math.abs(beta / SE);
    
    // 计算p值（精确计算）
    const pValue = 2 * (1 - normSDist(zScore));
    
    return {
        lower: lower.toFixed(4),
        upper: upper.toFixed(4),
        beta: beta.toFixed(4),
        SE: SE.toFixed(4),
        OR: OR.toFixed(4),
        pValue: pValue.toFixed(4),
        waldChi2: waldChi2.toFixed(4) // 新增 Wald χ² 值
    };
}

// 标准正态分布的累积分布函数
function normSDist(z) {
    let b1 = 0.319381530;
    let b2 = -0.356563782;
    let b3 = 1.781477937;
    let b4 = -1.821255978;
    let b5 = 1.330274429;
    let p = 0.2316419;
    let c = 0.39894228;

    if (z < 0) {
        z = -z;
    }

    let t = 1 / (1 + p * z);
    let prob = 1 - c * Math.exp(-z * z / 2) * (b1 * t + b2 * t * t + b3 * t * t * t + b4 * t * t * t * t + b5 * t * t * t * t * t);

    return prob;
}

// 使用与Cox回归相同的数字提取函数
function getPureNumberText(cell) {
    try {
        var text = cell.Range.Text.replace(/[^\d\.\-+eE]/g, "").trim();
        return isNaN(parseFloat(text)) ? null : text;
    } catch (e) {
        return null;
    }
}

// 使用与Cox回归完全相同的范围提取函数
function extractSingleCellRange(text) {
    // 1. 支持格式：1.2(0.8-1.6) 或 1.2（0.8～1.6）
    var match = text.match(/^([-+]?\d*\.?\d+)\s*[(（]\s*([-+]?\d*\.?\d+)\s*[—\-﹣–―~～\s]+\s*([-+]?\d*\.?\d+)\s*[)）]$/);
    if (match) return match[2] + " 到 " + match[3];
    
    // 2. 支持格式：0.8-1.6 或 0.8 ~ 1.697
    match = text.match(/^([-+]?\d*\.?\d+)\s*[—\-﹣–―~～]\s*([-+]?\d*\.?\d+)$/);
    if (match) return match[1] + " 到 " + match[2];
    
    // 3. 支持格式：0.209~1.697
    match = text.match(/^([-+]?\d*\.?\d+)\s*~\s*([-+]?\d*\.?\d+)$/);
    if (match) return match[1] + " 到 " + match[2];
    
    // 4. 支持格式：(0.8-1.6) 或 （0.8～1.6）
    match = text.match(/^[(（]\s*([-+]?\d*\.?\d+)\s*[—\-﹣–―~～\s]+\s*([-+]?\d*\.?\d+)\s*[)）]$/);
    if (match) return match[1] + " 到 " + match[2];
    
    return null;
}