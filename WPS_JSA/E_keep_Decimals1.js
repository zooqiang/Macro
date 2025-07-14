
function E_keep_Decimals() {
	

    var app = Application;
    var selection = app.Selection;
    
    // 检查是否在表格中
    if (!selection.Information(12)) {  // 12=wdWithInTable
        alert("请先在表格中选择要处理的单元格区域、行或列");
        return;
    }
    
    var table = selection.Tables(1);
    
    // 判断选择类型（行、列或任意区域）
    var selectionType = GetSelectionType(table, selection);
    if (!selectionType) {
        alert("请选择有效的区域（整行、整列或单元格区域）");
        return;
    }
    
    // 获取小数位数
    var decimalPlaces = InputBox("请输入要保留的小数位数（1-4）：", "设置小数位数", "2");
    if (decimalPlaces === false) return; // 用户点击取消
    
    decimalPlaces = parseInt(decimalPlaces);
    if (isNaN(decimalPlaces) || decimalPlaces < 1 || decimalPlaces > 4) {
        alert("请输入1到4之间的有效数字");
        return;
    }
    
    // 根据选择类型处理
    switch (selectionType.type) {
        case "column":
            // 处理整列
            for (var r = 1; r <= table.Rows.Count; r++) {
                try {
                    var cell = table.Cell(r, selectionType.number);
                    FormatCellNumber(cell.Range, decimalPlaces);
                } catch (e) {
                    // 跳过错误（如合并单元格）
                }
            }
            alert("列数值四舍五入完成！");
            break;
            
        case "row":
            // 处理整行
            for (var c = 1; c <= table.Columns.Count; c++) {
                try {
                    var cell = table.Cell(selectionType.number, c);
                    FormatCellNumber(cell.Range, decimalPlaces);
                } catch (e) {
                    // 跳过错误（如合并单元格）
                }
            }
            alert("行数值四舍五入完成！");
            break;
            
        case "range":
            // 处理任意选择区域
            var processedCells = 0;
            for (var r = selectionType.startRow; r <= selectionType.endRow; r++) {
                for (var c = selectionType.startCol; c <= selectionType.endCol; c++) {
                    try {
                        var cell = table.Cell(r, c);
                        if (FormatCellNumber(cell.Range, decimalPlaces)) {
                            processedCells++;
                        }
                    } catch (e) {
                        // 跳过错误（如合并单元格）
                    }
                }
            }
            alert("已处理 " + processedCells + " 个单元格的数值四舍五入！");
            break;
    }
}

// 改进后的选择类型判断函数
function GetSelectionType(table, selection) {
    // 首先尝试通过Information方法获取精确的行列信息
    try {
        var startRow = selection.Information(6); // wdStartOfRangeRowNumber
        var endRow = selection.Information(7);
        var startCol = selection.Information(8); // wdStartOfRangeColumnNumber
        var endCol = selection.Information(9);
        
        // 确保获取到了有效的位置信息
        if (startRow > 0 && endRow > 0 && startCol > 0 && endCol > 0) {
            // 判断是否是整列选择
            if (startCol === endCol && (endRow - startRow + 1) >= table.Rows.Count * 0.8) {
                return { type: "column", number: startCol };
            }
            // 判断是否是整行选择
            else if (startRow === endRow && (endCol - startCol + 1) >= table.Columns.Count * 0.8) {
                return { type: "row", number: startRow };
            }
            // 否则作为普通区域处理
            else {
                return {
                    type: "range",
                    startRow: startRow,
                    endRow: endRow,
                    startCol: startCol,
                    endCol: endCol
                };
            }
        }
    } catch (e) {
        // 如果Information方法失败，使用备用判断方法
        console.log("使用备用选择判断方法");
    }
    
    // 备用方法：通过选择范围判断
    var firstCell = table.Cell(1, 1);
    var lastCell = table.Cell(table.Rows.Count, table.Columns.Count);
    var tableStart = firstCell.Range.Start;
    var tableEnd = lastCell.Range.End;
    var selStart = selection.Start;
    var selEnd = selection.End;
    
    // 检查是否选择了整列
    for (var c = 1; c <= table.Columns.Count; c++) {
        var colStart = table.Cell(1, c).Range.Start;
        var colEnd = table.Cell(table.Rows.Count, c).Range.End;
        
        if (selStart <= colStart && selEnd >= colEnd) {
            return { type: "column", number: c };
        }
    }
    
    // 检查是否选择了整行
    for (var r = 1; r <= table.Rows.Count; r++) {
        var rowStart = table.Cell(r, 1).Range.Start;
        var rowEnd = table.Cell(r, table.Columns.Count).Range.End;
        
        if (selStart <= rowStart && selEnd >= rowEnd) {
            return { type: "row", number: r };
        }
    }
    
    // 如果既不是整行也不是整列，尝试作为区域处理
    try {
        // 获取选择中的第一个和最后一个单元格
        var firstSelCell = selection.Cells(1);
        var lastSelCell = selection.Cells(selection.Cells.Count);
        
        return {
            type: "range",
            startRow: firstSelCell.RowIndex,
            endRow: lastSelCell.RowIndex,
            startCol: firstSelCell.ColumnIndex,
            endCol: lastSelCell.ColumnIndex
        };
    } catch (e) {
        return null;
    }
}

function FormatCellNumber(cellRange, decimalPlaces) {
    var originalText = cellRange.Text.trim();
    
    // 更精确的数字提取正则表达式
    var numberMatch = originalText.match(/-?\d+[,.]?\d*/);
    if (!numberMatch) return false; // 不是数值则跳过
    
    var numberStr = numberMatch[0].replace(',', '.'); // 统一小数点为点号
    var number = parseFloat(numberStr);
    if (isNaN(number)) return false;
    
    // 精确四舍五入计算
    var factor = Math.pow(10, decimalPlaces);
    var rounded = Math.round(number * factor) / factor;
    
    // 格式化为字符串，确保显示正确的小数位数
    var formatted = rounded.toFixed(decimalPlaces);
    
    // 保留原文本中的非数字部分
    if (numberMatch.index > 0 || numberMatch[0].length < originalText.length) {
        var prefix = originalText.substring(0, numberMatch.index);
        var suffix = originalText.substring(numberMatch.index + numberMatch[0].length);
        cellRange.Text = prefix + formatted + suffix;
    } else {
        cellRange.Text = formatted;
    }
    
    return true; // 表示成功处理了数值
}