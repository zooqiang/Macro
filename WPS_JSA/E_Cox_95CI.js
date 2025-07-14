Attribute Module_Name = "E_Cox_95CI"
function E_Cox_95CI() {
    var selection = Application.Selection;
    
    if (!selection || !selection.Cells || selection.Cells.Count === 0) {
        MsgBox("请选择至少一个单元格", jsOKOnly, "提示");
        return;
    }
    
    var rangeText = "";
    var coxResults = null;
    
    if (selection.Cells.Count === 2) {
        var cell1Text = getPureNumberText(selection.Cells.Item(1));
        var cell2Text = getPureNumberText(selection.Cells.Item(2));
        
        if (cell1Text === null || cell2Text === null) {
            MsgBox("单元格必须包含有效数字", jsOKOnly, "错误");
            return;
        }
        
        var lower = parseFloat(cell1Text);
        var upper = parseFloat(cell2Text);
        
        // 检查95%CI范围是否合理
        if (lower > upper) {
            MsgBox("95%CI范围错误：下限值大于上限值\n请检查输入的置信区间数据", jsOKOnly, "错误");
            return;
        }
        
        rangeText = lower + " 到 " + upper;
        coxResults = calculateCoxResults(lower, upper);
    } 
    else if (selection.Cells.Count === 1) {
        var text = selection.Text.trim();
        rangeText = extractSingleCellRange(text);
        
        if (!rangeText) {
            MsgBox("未识别到有效的数值范围格式，请先修改区间符号\n支持格式：0.8-1.6、0.8~1.6、(0.8-1.6)、1.2(0.8-1.6)", jsOKOnly, "错误");
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
            
            coxResults = calculateCoxResults(lower, upper);
        }
    }
    
    // 构建批注内容
    var commentContent = rangeText;
    if (coxResults) {
        if (coxResults.error) {
            commentContent += "\r\n【错误】：" + coxResults.error;
        } else {
            commentContent += "\r\n【Cox回归结果：】";
            commentContent += "\r\nBeta: " + coxResults.beta;
            commentContent += "\r\nSE: " + coxResults.SE;
            commentContent += "\r\nWald χ²: " + coxResults.waldChi2; // 新增 Wald χ² 值
            commentContent += "\r\nP值: " + coxResults.pValue;
            commentContent += "\r\nHR: " + coxResults.HR;

        }
    }
    
    // 添加批注（不显示弹窗确认）
    selection.Comments.Add(selection.Range, commentContent + "\n");
}

// Cox回归专用计算函数（新增 Wald χ² 计算）
function calculateCoxResults(lower, upper) {
    // 安全检查 - HR的95%CI不能≤0
    if (lower <= 0 || upper <= 0) {
        return { error: "HR的95%CI不能≤0" };
    }
    
    // 计算方法验证（标准Cox回归计算）：
    // 1. 取95%CI的对数
    const logLower = Math.log(lower);
    const logUpper = Math.log(upper);
    
    // 2. 计算beta系数（log(HR)）
    const beta = (logLower + logUpper) / 2;
    
    // 3. 计算标准误(SE)
    const SE = (logUpper - logLower) / 3.92; // 3.92 = 2*1.96
    
    // 4. 计算HR值
    const HR = Math.exp(beta);
    
    // 5. 计算 Wald χ² 值（新增）
    const waldChi2 = Math.pow(beta / SE, 2);
    
    // 6. 计算Z值和P值
    const zScore = Math.abs(beta / SE);
    const pValue = 2 * (1 - normSDist(zScore)); // 精确计算p值
    
    return {
        lower: lower.toFixed(4),
        upper: upper.toFixed(4),
        beta: beta.toFixed(4),
        SE: SE.toFixed(4),
        HR: HR.toFixed(4),
        zScore: zScore.toFixed(3),
        pValue: pValue.toFixed(4), // 保留4位小数
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

// 增强的数字提取函数
function getPureNumberText(cell) {
    try {
        var text = cell.Range.Text.replace(/[^\d\.\-+eE]/g, "").trim();
        return isNaN(parseFloat(text)) ? null : text;
    } catch (e) {
        return null;
    }
}

// 增强的范围提取函数（支持更多格式）
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