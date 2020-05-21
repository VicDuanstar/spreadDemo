window.onload = function() {
    let spread = new GC.Spread.Sheets.Workbook(_getElementById('ss'));
    spread.setSheetCount(4);
    let spreadNS = GC.Spread.Sheets;
    let tableStyle = GC.Spread.Sheets.Tables.TableThemes.medium2;
    spread.suspendPaint();
    spread.suspendCalcService(false);
    initSheet3(spread.getSheet(2), tableStyle, spreadNS, spread);
    initSheet4(spread.getSheet(3), tableStyle, spreadNS, spread);
    initSheet2(spread.getSheet(1), tableStyle, spreadNS, spread);
    initSheet1(spread.getSheet(0), tableStyle, spreadNS, spread);
    spread.resumeCalcService(false);
    spread.resumePaint();

};

function initSheet4(sheet, tableStyle, spreadNS, spread) {
    // sheet.name("Operating Expenses");
    // let source = dataSource.sourceSheetFourTable1;
    // let table = sheet.tables.addFromDataSource("OperatingExpenses", 3, 1, source, tableStyle);
    _initLastThreeSheet(sheet, spreadNS, spread);
    _customerTable(table);
    sheet.getCell(24, 1).text("Total Operating Expenses");
    sheet.setArray(4, 2, dataSource.sheetFourFormula, true);
}

function initSheet3(sheet, tableStyle, spreadNS, spread) {
    // sheet.name("Personnel Expenses");
    // let source = dataSource.sourceSheetThreeTable1;
    // let table = sheet.tables.addFromDataSource("PersonnelExpenses", 3, 1, source, tableStyle);
    _initLastThreeSheet(sheet, spreadNS, spread);

    _customerTable(table);
    sheet.getCell(7, 1).text("Total Personnel Expenses");
    sheet.setArray(4, 2, dataSource.sheetThreeFormula, true);
}

function initSheet2(sheet, tableStyle, spreadNS, spread) {
    sheet.name("Income");
    let source = dataSource.sourceSheetTwoTable1;
    let table = sheet.tables.addFromDataSource("Income", 3, 1, source, tableStyle);
    _initLastThreeSheet(sheet, spreadNS, spread);
    _customerTable(table);
    sheet.getCell(7, 1).text("Total Income");
    sheet.setArray(4, 2, dataSource.sheetTowFormula, true);
}





function initSheet1(sheet, tableStyle, spreadNS, spread) {
    // sheet.name("Monthly Budget Summary");
    // let source = dataSource.sourceSheetOneTable1;
    // let table = sheet.tables.addFromDataSource("Totals", 3, 1, source, tableStyle);
    _customerTable(table);

    // let source1 = dataSource.sourceSheetOneTable2;
    // let table1 = sheet.tables.addFromDataSource("Top5Expenses", 10, 1, source1, tableStyle);
    _customerTable(table1);

    _setCellText(sheet, 6, 1, "Balance (Income minus Expenses)");
    _setCellText(sheet, 16, 1, "Total");

    sheet.addSpan(1, 1, 1, 3);
    sheet.addSpan(1, 4, 1, 2);
    sheet.addSpan(8, 1, 1, 4);
    sheet.getCell(1, 4)
        .text('Date')
        .font('normal normal 14.7px Gill Sans MT')
        .foreColor("rgb(53, 90, 97)")
        .vAlign(GC.Spread.Sheets.VerticalAlign.center)
        .hAlign(GC.Spread.Sheets.HorizontalAlign.right);

    //11111111111111111111111111111111
    _setWidthHeight(sheet, spread);
    sheet.setColumnWidth(7, 72);
    sheet.getCell(1, 1)
        .text('MONTHLY BUDGET')
        .font('normal normal 48px Gill Sans MT')
        .foreColor("rgb(53, 90, 97)");
    sheet.getCell(9, 1).text('WHAT ARE MY TOP 5 HIGHEST OPERATING EXPENSES?')
        .font('normal normal 14.7px Gill Sans MT');
    sheet.setRowHeight(8, 500);
    sheet.getRange(0, 0, 2, 6).backColor("#FFF");
    sheet.getRange(8, 0, 1, 6).backColor("#FFF");
    sheet.setRowCount(200, spreadNS.SheetArea.viewport);
    sheet.setColumnCount(40, spreadNS.SheetArea.viewport);

    let styleOfBackColor = new spreadNS.Style();
    styleOfBackColor.backColor = "#F2F2F2";
    _setBackColor(sheet, 0, 5, styleOfBackColor, spreadNS);

    _setDataValidator(sheet, spreadNS, 0, 1, dataSource.companyName);
    _setDataValidator(sheet, spreadNS, 1, 1, dataSource.monthlyBudget);
    _setDataValidator(sheet, spreadNS, 1, 2, dataSource.data);
    _setDataValidator(sheet, spreadNS, 3, 1, dataSource.BudgetTotals);
    _setDataValidator(sheet, spreadNS, 3, 2, dataSource.Estimated);
    _setDataValidator(sheet, spreadNS, 3, 3, dataSource.Actual);
    _setDataValidator(sheet, spreadNS, 3, 4, dataSource.Difference);
    _setDataValidator(sheet, spreadNS, 9, 1, dataSource.Top5HighestOperatingExpenses);
    _setDataValidator(sheet, spreadNS, 10, 1, dataSource.Expense);
    _setDataValidator(sheet, spreadNS, 10, 2, dataSource.Amount);
    _setDataValidator(sheet, spreadNS, 10, 3, dataSource._OfExpenses);
    _setDataValidator(sheet, spreadNS, 10, 4, dataSource._15Reduction);
    spread.addCustomName('BUDGET_Title', '="Monthly Budget Summary"!$B$2', 0, 0, 'comment');
    sheet.setArray(4, 2, dataSource.sheetOneFormula1, true);
    sheet.setArray(11, 1, dataSource.sheetOneFormula2, true);


    _customerChart(sheet, spreadNS, 34, 240, 687, 495, "B4:D6", "chart1", "BUDGET OVERVIEW");
}


/**
 * 根据数据源，图表放置的位置创建一个chart
 * @param {*} sheet 创建chart的工作表
 * @param {*} spreadNS spreadJS的主名称空间
 * @param {*} horizontalStart 水平方向开始的地方
 * @param {*} verticalStart 垂直方向开始的地方
 * @param {*} horizontalLength 水平方向的延伸长度
 * @param {*} verticalLength 垂直方向的延伸长度
 * @param {*} dataSource 数据源
 * @param {*} chartName chart的名称
 * @param {*} titleName chart的标题
 */
function _customerChart(sheet, spreadNS, horizontalStart, verticalStart, horizontalLength, verticalLength, dataSource, chartName, titleName) {
    sheet.charts.add(chartName, spreadNS.Charts.ChartType.columnClustered, horizontalStart, verticalStart, horizontalLength, verticalLength, dataSource);
    let chart = sheet.charts.get(chartName);
    let title = chart.title();
    title.text = titleName;
    title.fontFamily = 'Gill Sans MT';
    title.fontSize = "30";
    title.color = "rgb(53, 90, 97)";
    chart.title(title);
    let axes = chart.axes();
    axes.primaryValue.majorGridLine.visible = false;
    axes.primaryValue.lineStyle.width = 0;
    axes.primaryCategory.lineStyle.width = 0;
    axes.primaryValue.majorUnit = 10000;
    chart.axes(axes);
    let chartArea = chart.chartArea();
    chartArea.border.width = 0;
    chart.chartArea(chartArea);
}

/**
 * 给特定的cell设置DataValidator
 * @param {*} sheet 当前的工作表
 * @param {*} spreadNS spreadJS的主命名空间
 * @param {*} row 该cell的行号
 * @param {*} colum 该cell的列号
 * @param {*} dataSource DataValidator的inputMessage
 */
function _setDataValidator(sheet, spreadNS, row, colum, dataSource) {
    let dataValidation = new spreadNS.DataValidation.createDateValidator("");
    dataValidation.inputMessage(dataSource);
    sheet.setDataValidator(row, colum, dataValidation);
}


/**
 * 设置sheet的row和Colum
 * @param {*} sheet     需要改变的工作表
 * @param {*} spread    工作簿
 */
function _setWidthHeight(sheet, spread) {
    sheet.options.gridline.showHorizontalGridline = false;
    sheet.options.gridline.showVerticalGridline = false;
    sheet.getCell(0, 1).text('COMPANY NAME').font('normal normal 21.3px Gill Sans MT').foreColor("rgb(53, 90, 97)");
    let dataSourceOfWH = _getDataSource(sheet, spread);
    let row = dataSourceOfWH.row;
    let colum = dataSourceOfWH.colum;

    _setRowHeight(sheet, row);
    _setColumWidth(sheet, colum);

}


function _setRowHeight(sheet, row) {
    let startColum = row[0];
    for (let i = 1; i < row.length; i++) {
        sheet.setRowHeight(startColum++, 22 + row[i]);
    }
}

function _setColumWidth(sheet, colum) {
    let startColum = colum[0];
    for (let i = 1; i < colum.length; i++) {
        sheet.setColumnWidth(startColum++, 72 + colum[i]);
    }
}


/**
 * 获取该工作表的数据源
 * @param {*} sheet 该工作表
 * @param {*} spread 该工作簿
 */
function _getDataSource(sheet, spread) {
    let dataSourceOfWH = "";
    switch (spread.getSheetIndex(sheet.name())) {
        case 0:
            dataSourceOfWH = dataSource.sheetOneScare;
            break;
        case 1:
            dataSourceOfWH = dataSource.sheetTwoScare;
            break;
        case 2:
            dataSourceOfWH = dataSource.sheetThreeScare;
            break;
        case 3:
            dataSourceOfWH = dataSource.sheetFourScare;
            break;
        default:
            break;
    }
    return dataSourceOfWH;
}

function _initLastThreeSheet(sheet, spreadNS, spread) {
    sheet.setFormula(1, 1, "BUDGET_Title");
    sheet.getCell(1, 1)
        .font('normal normal 48px Gill Sans MT')
        .foreColor("rgb(53, 90, 97)");
    _setWidthHeight(sheet, spread);
    sheet.setColumnWidth(4, 0);
    sheet.setColumnWidth(5, 152);
    sheet.getRange(0, 0, 2, 7).backColor("#FFF");
    let styleOfBackColor = new spreadNS.Style();
    styleOfBackColor.backColor = "#DCEBEA";
    _setBackColor(sheet, 0, 6, styleOfBackColor, spreadNS);
}

function _customerTable(table) {
    table.showFooter(true);
    table.useFooterDropDownList(true);
    table.filterButtonVisible(false);
}



function _setCellText(sheet, row, colum, text) {
    sheet.getCell(row, colum).text(text);
}

/**
 * 在开始列和结束列之间设置背景
 * @param {*} sheet 需要设置的工作表
 * @param {*} columStart 开始的列号
 * @param {*} columEnd 结束的列
 * @param {*} style 需要设置的样式
 * @param {*} spreadNS 工作表的上一层(包含所有的工作表的共同属性及动作)
 */
function _setBackColor(sheet, columStart, columEnd, style, spreadNS) {
    for (let i = columStart; i <= columEnd; i++) {
        sheet.setStyle(-1, i, style, spreadNS.SheetArea.viewport);
    }
}

/**
 * 获取DOC的一个节点(根据节点id)
 * @param {String} id 节点的id值
 */
function _getElementById(id) {
    return document.getElementById(id);
}