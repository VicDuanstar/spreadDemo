class Sheet {
    dataOfSheet = "";
    constructor(sheetNumber, spreadNS, spread, dataSource) {
        this.spread = spread;
        this.sheetNumber = sheetNumber;
        this.spreadNS = spreadNS;
        this.dataSource = dataSource;
        this.initSheet();
    }

    findDataOfSheet(sheetNumber, dataSource) {
        return dataSource.sheet[sheetNumber];
    }
    initSheet() {
        let sheetNumber = this.sheetNumber;
        let sheet = this.spread.getSheet(sheetNumber);

        let dataSource = this.findDataOfSheet(sheetNumber, this.dataSource);
        let spread = this.spread;
        let spreadNS = this.spreadNS;
        sheet.name(dataSource.name);
        this.addTable(sheet, dataSource);

        this._setCellText(dataSource.cellText, sheet);


        this._mergeCell(dataSource.mergeCell, sheet);


        this._setCellStyle(sheet);

        this._setWidthHeight(sheet, dataSource);

        this._setNumberOfRowAndColum(sheet, spreadNS, dataSource);

        let styleOfBackColor = new spreadNS.Style();
        this._setBackColor(sheet, dataSource.backColor, spreadNS, styleOfBackColor);

        this._setRangeColor(sheet, dataSource.rangeColor);


        this._setDataValidator(sheet, spreadNS, dataSource.dataValidation);


        this.addChart(sheet, spreadNS, dataSource.chart);

        this._setFormatter(sheet, dataSource.formatter);
    }

    addTable(sheet, dataSource) {
        if (dataSource.table.length <= 0) {
            return;
        } else {
            for (let i = 0; i < dataSource.table.length; i++) {
                new myTable(sheet, dataSource.table[i]);
            }
        }
    }

    addChart(sheet, spreadNS, dataSource) {
        if (typeof dataSource === 'undefined' || dataSource.length <= 0) {
            return;
        }
        for (let i = 0; i < dataSource.length; i++) {
            let chart = dataSource[i];
            new myChart(sheet, spreadNS, chart);
        }
    }

    _setCellText(cellText, sheet) {
        if (typeof cellText === 'undefined' || cellText.length <= 0) {
            return;
        }
        for (let i = 0; i < cellText.length; i++) {
            sheet.getCell(cellText[i].rowNumber, cellText[i].columNumber).text(cellText[i].text);
        }
    }

    _mergeCell(mergeCell, sheet) {
        if (typeof mergeCell === 'undefined' || mergeCell.length <= 0) {
            return;
        }
        for (let i = 0; i < mergeCell.length;) {
            sheet.addSpan(mergeCell[i++], mergeCell[i++], mergeCell[i++], mergeCell[i++]);
        }
    }
    _setCellStyle(sheet) {
        sheet.getCell(1, 4)
            .font('normal normal 14.7px Gill Sans MT')
            .foreColor("rgb(53, 90, 97)")
            .vAlign(GC.Spread.Sheets.VerticalAlign.center)
            .hAlign(GC.Spread.Sheets.HorizontalAlign.right);
        sheet.getCell(1, 1)
            .font('normal normal 48px Gill Sans MT')
            .foreColor("rgb(53, 90, 97)");
        sheet.getCell(9, 1)
            .font('normal normal 14.7px Gill Sans MT');


    }

    _setWidthHeight(sheet, dataSource) {
        sheet.options.gridline.showHorizontalGridline = false;
        sheet.options.gridline.showVerticalGridline = false;
        sheet.getCell(0, 1).font('normal normal 21.3px Gill Sans MT').foreColor("rgb(53, 90, 97)");
        let dataSourceOfWH = dataSource.sheetScare;
        let row = dataSourceOfWH.row;
        let colum = dataSourceOfWH.colum;
        this._setRowHeight(sheet, row);
        this._setColumWidth(sheet, colum);

    }

    _setRowHeight(sheet, row) {
        let startColum = row[0];
        for (let i = 1; i < row.length; i++) {
            sheet.setRowHeight(startColum++, 22 + row[i]);
        }
    }

    _setColumWidth(sheet, colum) {
        let startColum = colum[0];
        for (let i = 1; i < colum.length; i++) {
            sheet.setColumnWidth(startColum++, 72 + colum[i]);
        }
    }
    _setNumberOfRowAndColum(sheet, spreadNS, dataSource) {
        sheet.setRowCount(dataSource.rowCount, spreadNS.SheetArea.viewport);
        sheet.setColumnCount(dataSource.columCount, spreadNS.SheetArea.viewport);
    }
    _setBackColor(sheet, dataSource, spreadNS, styleOfBackColor) {
        if (typeof dataSource === 'undefined' || dataSource.length <= 0) {
            return;
        }

        for (let n = 0; n < dataSource.length; n++) {
            styleOfBackColor.backColor = dataSource[n].style;
            for (let i = dataSource[n].columStart; i <= dataSource[n].columEnd; i++) {
                sheet.setStyle(-1, i, styleOfBackColor, spreadNS.SheetArea.viewport);
            }
        }
    }

    _setRangeColor(sheet, dataSource) {
        if (typeof dataSource === 'undefined' || dataSource.length <= 0) {
            return;
        }
        for (let i = 0; i < dataSource.length; i++) {
            sheet.getRange(dataSource[i].rowStart, dataSource[i].columStart, dataSource[i].rowLength, dataSource[i].columLength).backColor(dataSource[i].style);
        }
    }

    _setDataValidator(sheet, spreadNS, dataSource) {
        if (typeof dataSource === 'undefined' || dataSource.length <= 0) {
            return;
        }
        let dataValidation = "";
        for (let i = 0; i < dataSource.length; i++) {
            dataValidation = new spreadNS.DataValidation.createDateValidator("");
            dataValidation.inputMessage(dataSource[i].text);
            sheet.setDataValidator(dataSource[i].rowNumber, dataSource[i].columNumber, dataValidation);
        }

    }

    _setFormatter(sheet, dataSource) {
        if (typeof dataSource === 'undefined' || dataSource.length <= 0) {
            return;
        }
        for (let i = 0; i < dataSource.length; i++) {
            sheet.getRange(dataSource[i].rowStart, dataSource[i].columStart, dataSource[i].rowLength, dataSource[i].columLength).formatter(dataSource[i].formatter);
        }
    }
}