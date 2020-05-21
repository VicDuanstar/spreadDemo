class myTable {
    constructor(sheet, dataOfTable) {
        this.sheet = sheet;
        this.dataOfTable = dataOfTable;
        this.initTable();
    }
    initTable() {
        let sheet = this.sheet;
        let dataSource = this.dataOfTable;
        let table = this._setTable(sheet, dataSource);
    }

    _customerTable(table) {
        table.showFooter(true);
        table.useFooterDropDownList(true);
        table.filterButtonVisible(false);
    }

    _setTable(sheet, dataSource) {
        if (typeof dataSource === 'undefined') {
            return;
        }
        let tableTheme = new GC.Spread.Sheets.Tables.TableTheme();

        this._setStyle(dataSource.style, tableTheme);
        let table = sheet.tables.addFromDataSource(dataSource.name, dataSource.horizontalLocation, dataSource.verticalLocal, dataSource.dataOfTable);
        // table.bandRows(false);
        table.style(tableTheme);
        table.highlightLastColumn(true);
        this._customerTable(table);
        this._setFormula(sheet, dataSource);
        return table;
    }

    _setFormula(sheet, dataSource) {
        if (typeof dataSource.tableFormula === 'undefined' || typeof dataSource.FormulaRange === 'undefined') {
            return;
        }
        let count = 0;
        let formulaRange = dataSource.FormulaRange;
        let tableFormula = dataSource.tableFormula;
        for (let i = 0; i < formulaRange.length;) {
            sheet.getRange(formulaRange[i++], formulaRange[i++], formulaRange[i++], formulaRange[i++]).formula(tableFormula[count++]);
        }
    }
    _setStyle(dataSource, tableTheme) {
        if (typeof dataSource === 'undefined') {
            return;
        }

        let tableStyle = new GC.Spread.Sheets.Tables.TableStyle();
        tableStyle.backColor = dataSource[0].dataColor;
        tableTheme.wholeTableStyle(tableStyle);

        tableStyle = new GC.Spread.Sheets.Tables.TableStyle();
        tableStyle.backColor = dataSource[0].headerRowColor;
        tableTheme.headerRowStyle(tableStyle);

        tableStyle = new GC.Spread.Sheets.Tables.TableStyle();
        tableStyle.backColor = dataSource[0].lastRowColor;
        tableTheme.footerRowStyle(tableStyle);

        tableStyle = new GC.Spread.Sheets.Tables.TableStyle();
        tableStyle.backColor = dataSource[0].lastColumColor;
        tableTheme.highlightLastColumnStyle(tableStyle);


    }

}