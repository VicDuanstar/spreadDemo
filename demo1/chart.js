class myChart {
    constructor(sheet, spreadNS, dataSource) {
        this.sheet = sheet;
        this.spreadNS = spreadNS;
        this.dataSource = dataSource;
        this.initChart();
    }


    initChart() {
        let sheet = this.sheet;
        let dataSource = this.dataSource;
        sheet.charts.add(dataSource.name, dataSource.chartType, dataSource.horizontalStart, dataSource.verticalStart, dataSource.horizontalLength, dataSource.verticalLength, dataSource.chartData);
        let chart = sheet.charts.get(dataSource.name);
        let title = chart.title();
        title.text = dataSource.title;
        title.fontFamily = 'Gill Sans MT';
        title.fontSize = "21";
        title.color = "rgb(53, 90, 97)";

        let axes = chart.axes();
        axes.primaryValue.majorGridLine.visible = false;
        axes.primaryValue.lineStyle.width = 0;
        axes.primaryCategory.lineStyle.width = 0;
        axes.primaryValue.majorUnit = 10000;
        chart.axes(axes);

        let chartArea = chart.chartArea();
        chartArea.border.width = 0;
        // chartArea.backColorTransparency = 0.9;
        chartArea.fontSize = 16;
        chart.chartArea(chartArea);

        let legend = chart.legend();
        legend.layout = {
                x: 0.37,
                y: 0,
                width: 0,
                height: 0
            }
            // legend.showLegendWithoutOverlapping(true);
        chart.legend(legend);

        let series = chart.series();
        let series1 = series.get(0);
        let series2 = series.get(1);
        series1.backColor = "#E2E8AC";
        series2.backColor = "#559592";
        series.set(0, series1);
        series.set(1, series2);
        chart.series(series);

        chart.title(title);
    }
}