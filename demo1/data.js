let dataSource = {
    length: 4,

    sheet: [

        {
            name: "Monthly Budget Summary",
            rowCount: 200,
            columCount: 40,
            mergeCell: [1, 1, 1, 3, 1, 4, 1, 2, 8, 1, 1, 4],
            formatter: [
                { rowStart: 4, columStart: 2, rowLength: 3, columLength: 3, formatter: "#,##0.00_);[Red](#,##0.00)" },
                { rowStart: 11, columStart: 2, rowLength: 6, columLength: 3, formatter: "#,##0.00_);[Red](#,##0.00)" },
                { rowStart: 11, columStart: 3, rowLength: 6, columLength: 1, formatter: "#.#%" },
            ],
            cellText: [
                { rowNumber: 6, columNumber: 1, text: "Balance (Income minus Expenses)" },
                { rowNumber: 16, columNumber: 1, text: "Total" },
                { rowNumber: 1, columNumber: 4, text: "Date" },
                { rowNumber: 1, columNumber: 1, text: "MONTHLY BUDGET" },
                { rowNumber: 9, columNumber: 1, text: "WHAT ARE MY TOP 5 HIGHEST OPERATING EXPENSES?" },
                { rowNumber: 0, columNumber: 1, text: "COMPANY NAME" },
            ],

            backColor: [
                { columStart: 0, columEnd: 10, style: "#F2F2F2" },
            ],

            rangeColor: [
                { rowStart: 0, columStart: 0, rowLength: 2, columLength: 7, style: "#FFF" },
            ],

            dataValidation: [
                { rowNumber: 0, columNumber: 1, text: "Enter Company Name in this cell", },
                { rowNumber: 1, columNumber: 1, text: "Title of this worksheet is in this cell. Enter Date in cell at right. Budget Totals are automatically calculated in Totals table starting in cell B4", },
                { rowNumber: 1, columNumber: 2, text: "Enter Date in this cell. Budget overview chart is in cell B9", },
                { rowNumber: 3, columNumber: 1, text: "Budget Totals for Income & Expenses, both estimated & actual, are automatically calculated from amounts entered in other worksheets. Balance & Difference are automatically adjusted", },
                { rowNumber: 3, columNumber: 2, text: "Estimated totals are automatically calculated in this column under this heading", },
                { rowNumber: 3, columNumber: 3, text: "Actual totals are automatically calculated in this column under this heading", },
                { rowNumber: 3, columNumber: 4, text: "Difference of Estimated and Actual Totals is automatically calculated in this column under this heading", },
                { rowNumber: 9, columNumber: 1, text: "Top 5 Operating Expenses are automatically updated in table below", },
                { rowNumber: 10, columNumber: 1, text: "Top 5 Expense items are automatically updated in this column under this heading", },
                { rowNumber: 10, columNumber: 2, text: "Amount is automatically updated in this column under this heading", },
                { rowNumber: 10, columNumber: 3, text: "Percent of Expenses is automatically calculated in this column under this heading", },
                { rowNumber: 10, columNumber: 4, text: "15 percent Reduction amount is automatically calculated in this column under this heading", },
            ],
            table: [{
                    name: "Totals",
                    horizontalLocation: 3,
                    verticalLocal: 1,
                    tableStyle: GC.Spread.Sheets.Tables.TableThemes.medium2,
                    dataOfTable: [
                        { 'BUDGET TOTALS': "Income", ESTIMATED: 23, ACTUAL: 21, DIFFERENCE: 21 },
                        { 'BUDGET TOTALS': "Expenses", ESTIMATED: 23, ACTUAL: 21, DIFFERENCE: 21 },
                    ],

                    style: [{
                        headerRowColor: "#B1C9B3",
                        firstColumColor: "",
                        lastRowColor: "#F0F3D6",
                        lastColumColor: "#F0F3D6",
                        dataColor: "#E9E9E2",
                        formatter: "#,##0.00_);[Red](#,##0.00)",
                    }],


                    tableFormula: [
                        '=Income[[#Totals],[ESTIMATED]]',
                        '=Income[[#Totals],[ACTUAL]]',
                        '=[@ACTUAL]-[@ESTIMATED]',
                        '=OperatingExpenses[[#Totals],[ESTIMATED]]+PersonnelExpenses[[#Totals],[ESTIMATED]]',
                        '=OperatingExpenses[[#Totals],[ACTUAL]]+PersonnelExpenses[[#Totals],[ACTUAL]]',
                        '=[@ESTIMATED]-[@ACTUAL]',
                        '=C5-C6',
                        '=D5-D6',
                        '=Totals[[#Totals],[ACTUAL]]-Totals[[#Totals],[ESTIMATED]]'
                    ],
                    FormulaRange: [4, 2, 1, 1, 4, 3, 1, 1, 4, 4, 1, 1, 5, 2, 1, 1, 5, 3, 1, 1, 5, 4, 1, 1, 6, 2, 1, 1, 6, 3, 1, 1, 6, 4, 1, 1],
                },

                {
                    name: "Top5Expenses",
                    horizontalLocation: 10,
                    verticalLocal: 1,
                    tableStyle: GC.Spread.Sheets.Tables.TableThemes.medium4,
                    dataOfTable: [
                        { EXPENSE: "Total", AMOUNT: 0, "% OF EXPENSES": 0, "15% REDUCTION": 0 },
                        { EXPENSE: "Total", AMOUNT: 0, "% OF EXPENSES": 0, "15% REDUCTION": 0 },
                        { EXPENSE: "Total", AMOUNT: 0, "% OF EXPENSES": 0, "15% REDUCTION": 0 },
                        { EXPENSE: "Total", AMOUNT: 0, "% OF EXPENSES": 0, "15% REDUCTION": 0 },
                        { EXPENSE: "Total", AMOUNT: 0, "% OF EXPENSES": 0, "15% REDUCTION": 0 },
                    ],
                    style: [{
                        headerRowColor: "#B1C9B3",
                        firstColumColor: "",
                        lastRowColor: "#F0F3D6",
                        lastColumColor: "#F0F3D6",
                        dataColor: "#E9E9E2",
                        formatter: "#,##0.00_);[Red](#,##0.00)",
                    }],

                    tableFormula: [
                        '=INDEX(OperatingExpenses,MATCH([@AMOUNT],OperatingExpenses[TOP 5 AMOUNT],0),1)',
                        '=LARGE(OperatingExpenses[TOP 5 AMOUNT],1)',
                        '=LARGE(OperatingExpenses[TOP 5 AMOUNT],2)',
                        '=LARGE(OperatingExpenses[TOP 5 AMOUNT],3)',
                        '=LARGE(OperatingExpenses[TOP 5 AMOUNT],4)',
                        '=LARGE(OperatingExpenses[TOP 5 AMOUNT],5)',
                        '=[@AMOUNT]/$D$6',
                        '=[@AMOUNT]*0.15',
                        '=SUBTOTAL(109,[AMOUNT])',
                        '=SUBTOTAL(109,[% OF EXPENSES])',
                        '=SUBTOTAL(109,[15% REDUCTION])',
                    ],
                    FormulaRange: [11, 1, 5, 1, 11, 2, 1, 1, 12, 2, 1, 1, 13, 2, 1, 1, 14, 2, 1, 1, 15, 2, 1, 1, 11, 3, 5, 1, 11, 4, 5, 1, 16, 2, 1, 1, 16, 3, 1, 1, 16, 4, 1, 1],
                },
            ],
            chart: [{
                name: "chart1",
                chartData: "B4:D6",
                title: "BUDGET OVERVIEW",
                horizontalStart: 34,
                verticalStart: 240,
                horizontalLength: 687,
                verticalLength: 440,
                chartType: GC.Spread.Sheets.Charts.ChartType.columnClustered,
            }],
            sheetScare: { row: [0, 20, 33, -2, 7, 1, 1, 1, 0, 425, 0, 7], colum: [0, -39, 162, 80, 80, 80, -39, -39] },

        },

        {
            name: "Income",
            rowCount: 200,
            columCount: 40,
            formatter: [
                { rowStart: 4, columStart: 2, rowLength: 4, columLength: 4, formatter: "#,##0.00_);[Red](#,##0.00)" },
            ],

            cellText: [
                { rowNumber: 1, columNumber: 1, text: "MONTHLY BUDGET" },
                { rowNumber: 0, columNumber: 1, text: "COMPANY NAME" },
                { rowNumber: 7, columNumber: 1, text: "Total Income" }
            ],
            backColor: [
                { columStart: 0, columEnd: 6, style: "#DCEBEA" }
                // { columStart: 0, columEnd: 6, style: "green" }
            ],
            rangeColor: [
                { rowStart: 0, columStart: 0, rowLength: 2, columLength: 7, style: "#FFF" },
            ],
            table: [{
                name: "Income",
                horizontalLocation: 3,
                verticalLocal: 1,
                tableStyle: GC.Spread.Sheets.Tables.TableThemes.medium3,
                dataOfTable: [
                    { INCOME: "Net sales", ESTIMATED: 60000.00, ACTUAL: 54000.00, 'TOP 5 AMOUNT': '?', DIFFERENCE: "?" },
                    { INCOME: "Interest income", ESTIMATED: 3000.00, ACTUAL: 3000.00, 'TOP 5 AMOUNT': '?', DIFFERENCE: "?" },
                    { INCOME: "Asset sales (gain/loss)", ESTIMATED: 300.00, ACTUAL: 450.00, 'TOP 5 AMOUNT': '?', DIFFERENCE: "?" },
                ],
                style: [{
                    headerRowColor: "#B1C9B3",
                    firstColumColor: "",
                    lastRowColor: "#F0F3D6",
                    lastColumColor: "#F0F3D6",
                    dataColor: "#E9E9E2",
                }],
                dataValidation: [
                    { rowNumber: 1, columNumber: 2, text: "15 percent Reduction amount is automatically calculated in this column under this heading", },
                    { rowNumber: 1, columNumber: 2, text: "Enter Income details in this column under this heading. Use heading filters to find specific entries", },
                    { rowNumber: 1, columNumber: 2, text: "Enter Estimated amount in this column under this heading", },
                    { rowNumber: 1, columNumber: 2, text: "Enter Actual amount in this column under this heading", },
                ],
                tableFormula: [
                    '=[@ACTUAL]+(10^-6)*ROW([@ACTUAL])',
                    '=[@ACTUAL]-[@ESTIMATED]',
                    '=SUBTOTAL(109,[ESTIMATED])',
                    '=SUBTOTAL(109,[ACTUAL])',
                    '=SUBTOTAL(109,[DIFFERENCE])'
                ],
                FormulaRange: [4, 4, 3, 1, 4, 5, 3, 1, 7, 2, 1, 1, 7, 3, 1, 1, 7, 5, 1, 1],
            }, ],
            sheetScare: { row: [0, 20, 33, -2], colum: [0, -39, 162, 80, 80, 0, 80, -39, -39] },

        },

        {
            name: "Personnel Expenses",
            rowCount: 200,
            columCount: 40,
            formatter: [
                { rowStart: 4, columStart: 2, rowLength: 4, columLength: 4, formatter: "#,##0.00_);[Red](#,##0.00)" },
            ],
            cellText: [
                { rowNumber: 1, columNumber: 1, text: "MONTHLY BUDGET" },
                { rowNumber: 0, columNumber: 1, text: "COMPANY NAME" },
                { rowNumber: 7, columNumber: 1, text: "Total Personnel Expenses" }
            ],
            backColor: [
                { columStart: 0, columEnd: 6, style: "#DCEBEA" }
            ],
            rangeColor: [
                { rowStart: 0, columStart: 0, rowLength: 2, columLength: 7, style: "#FFF" },
            ],
            table: [{
                    name: "PersonnelExpenses",
                    horizontalLocation: 3,
                    verticalLocal: 1,
                    tableStyle: GC.Spread.Sheets.Tables.TableThemes.medium1,
                    dataOfTable: [
                        { INCOME: "Net sales", ESTIMATED: 9500.00, ACTUAL: 9600.00, 'TOP 5 AMOUNT': '?', DIFFERENCE: "?" },
                        { INCOME: "Interest income", ESTIMATED: 4000.00, ACTUAL: 0.00, 'TOP 5 AMOUNT': '?', DIFFERENCE: "?" },
                        { INCOME: "Asset sales (gain/loss)", ESTIMATED: 5000.00, ACTUAL: 4500.00, 'TOP 5 AMOUNT': '?', DIFFERENCE: "?" },
                    ],
                    style: [{
                        headerRowColor: "#B1C9B3",
                        firstColumColor: "",
                        lastRowColor: "#F0F3D6",
                        lastColumColor: "#F0F3D6",
                        dataColor: "#E9E9E2",
                        formatter: "#,##0.00_);[Red](#,##0.00)",
                    }],
                    dataValidation: [
                        { rowNumber: 1, columNumber: 2, text: "Enter Personnel Expenses in this column under this heading. Use heading filters to find specific entries", },
                        { rowNumber: 1, columNumber: 2, text: "Difference of Estimated and Actual Personnel Expenses is automatically calculated in this column under this heading", },
                    ],
                    tableFormula: [
                        '=[@ACTUAL]+(10^-6)*ROW([@ACTUAL])',
                        '=[@ESTIMATED]-[@ACTUAL]',
                        '=SUBTOTAL(109,[ESTIMATED])',
                        '=SUBTOTAL(109,[ACTUAL])',
                        '=SUBTOTAL(109,[DIFFERENCE])'
                    ],
                    FormulaRange: [4, 4, 3, 1, 4, 5, 3, 1, 7, 2, 1, 1, 7, 3, 1, 1, 7, 5, 1, 1],

                },

            ],
            sheetScare: { row: [0, 20, 33, -2], colum: [0, -39, 162, 80, 80, 0, 80, -39, -39] },
        },

        {
            name: "Operating Expenses",
            rowCount: 200,
            columCount: 40,
            formatter: [
                { rowStart: 4, columStart: 2, rowLength: 21, columLength: 4, formatter: "#,##0.00_);[Red](#,##0.00)" },
            ],
            cellText: [
                { rowNumber: 1, columNumber: 1, text: "MONTHLY BUDGET" },
                { rowNumber: 0, columNumber: 1, text: "COMPANY NAME" },
                { rowNumber: 24, columNumber: 1, text: "Total Operating Expenses" }
            ],
            backColor: [
                { columStart: 0, columEnd: 6, style: "#DCEBEA" }
            ],
            rangeColor: [
                { rowStart: 0, columStart: 0, rowLength: 2, columLength: 7, style: "#FFF" },
            ],
            table: [{

                name: "OperatingExpenses",
                horizontalLocation: 3,
                verticalLocal: 1,
                tableStyle: GC.Spread.Sheets.Tables.TableThemes.medium2,

                style: [{
                    headerRowColor: "#B1C9B3",
                    firstColumColor: "",
                    lastRowColor: "#F0F3D6",
                    lastColumColor: "#F0F3D6",
                    dataColor: "#E9E9E2",
                }],
                dataOfTable: [
                    { OPERATING_EXPENSES: "Advertising", ESTIMATED: 3000.00, ACTUAL: 2500.00, "TOP 5 AMOUNT": 2500.00, DIFFERENCE: 500.00 },
                    { OPERATING_EXPENSES: "Bad debts", ESTIMATED: 2000.00, ACTUAL: 2000.00, "TOP 5 AMOUNT": "	2,000.00", DIFFERENCE: "	0.00 " },
                    { OPERATING_EXPENSES: "Cash discounts", ESTIMATED: 1500.00, ACTUAL: 2175.00, "TOP 5 AMOUNT": "	2,175.00", DIFFERENCE: "	(675.00)" },
                    { OPERATING_EXPENSES: "Delivery costs", ESTIMATED: 2000.00, ACTUAL: 1500.00, "TOP 5 AMOUNT": "	1,500.00", DIFFERENCE: "	500.00 " },
                    { OPERATING_EXPENSES: "Depreciation", ESTIMATED: 1000.00, ACTUAL: 1000.00, "TOP 5 AMOUNT": "	1,000.00", DIFFERENCE: "	0.00 " },
                    { OPERATING_EXPENSES: "Dues and subscriptions", ESTIMATED: 500.00, ACTUAL: 525.00, "TOP 5 AMOUNT": "   525.00 ", DIFFERENCE: "		(25.00)" },
                    { OPERATING_EXPENSES: "Insurance", ESTIMATED: 1300.00, ACTUAL: 1275.00, "TOP 5 AMOUNT": "	1,275.00", DIFFERENCE: "	25.00 " },
                    { OPERATING_EXPENSES: "Interest", ESTIMATED: 2000.00, ACTUAL: 2200.00, "TOP 5 AMOUNT": "	2,200.00", DIFFERENCE: "	(200.00)" },
                    { OPERATING_EXPENSES: "Legal and auditing", ESTIMATED: 1000.00, ACTUAL: 800.00, "TOP 5 AMOUNT": "	800.00 	", DIFFERENCE: "	200.00 " },
                    { OPERATING_EXPENSES: "Maintenance and repairs", ESTIMATED: 4500.00, ACTUAL: 4600.00, "TOP 5 AMOUNT": "	4,600.00", DIFFERENCE: "	(100.00)" },
                    { OPERATING_EXPENSES: "Office supplies", ESTIMATED: 800.00, ACTUAL: 750.00, "TOP 5 AMOUNT": "  750.00 	", DIFFERENCE: "	50.00 " },
                    { OPERATING_EXPENSES: "Postage", ESTIMATED: 400.00, ACTUAL: 350.00, "TOP 5 AMOUNT": "  350.00 	", DIFFERENCE: "	50.00 " },
                    { OPERATING_EXPENSES: "Rent or mortgage", ESTIMATED: 4100.00, ACTUAL: 4500.00, "TOP 5 AMOUNT": "	4,500.00", DIFFERENCE: "	(400.00)" },
                    { OPERATING_EXPENSES: "Sales expenses", ESTIMATED: 350.00, ACTUAL: 400.00, "TOP 5 AMOUNT": "  400.00 	", DIFFERENCE: "	(50.00)" },
                    { OPERATING_EXPENSES: "Shipping and storage", ESTIMATED: 900.00, ACTUAL: 840.00, "TOP 5 AMOUNT": "  840.00 	", DIFFERENCE: "	60.00 " },
                    { OPERATING_EXPENSES: "Supplies", ESTIMATED: 5000.00, ACTUAL: 4500.00, "TOP 5 AMOUNT": "	4,500.00", DIFFERENCE: "	500.00 " },
                    { OPERATING_EXPENSES: "Taxes", ESTIMATED: 3000.00, ACTUAL: 3200.00, "TOP 5 AMOUNT": "	3,200.00", DIFFERENCE: "	(200.00)" },
                    { OPERATING_EXPENSES: "Telephone", ESTIMATED: 250.00, ACTUAL: 280.00, "TOP 5 AMOUNT": "  280.00 	", DIFFERENCE: "	(30.00)" },
                    { OPERATING_EXPENSES: "Utilities", ESTIMATED: 1400.00, ACTUAL: 1385.00, "TOP 5 AMOUNT": "	1,385.00", DIFFERENCE: "	15.00 " },
                    { OPERATING_EXPENSES: "Other", ESTIMATED: 1000.00, ACTUAL: 750.00, "TOP 5 AMOUNT": "	750.00 	", DIFFERENCE: "	250.00 " },
                ],
                dataValidation: [
                    { rowNumber: 1, columNumber: 2, text: "Enter Operating Expenses in this column under this heading. Use heading filters to find specific entries", },
                    { rowNumber: 1, columNumber: 2, text: "Difference of Estimated and Actual Operating Expenses is automatically calculated in this column under this heading", },
                ],
                tableFormula: [
                    '=[@ACTUAL]+(10^-6)*ROW([@ACTUAL])',
                    '=[@ESTIMATED]-[@ACTUAL]',
                    '=SUBTOTAL(109,[ESTIMATED])',
                    '=SUBTOTAL(109,[ACTUAL])',
                    '=SUBTOTAL(109,[DIFFERENCE])'
                ],
                FormulaRange: [4, 4, 20, 1, 4, 5, 20, 1, 24, 2, 1, 1, 24, 3, 1, 1, 24, 5, 1, 1],
            }, ],

            sheetScare: { row: [0, 20, 33, -2], colum: [0, -39, 162, 80, 80, 0, 80, -39, -39] }
        }
    ]
}