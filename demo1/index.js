// import Sheet from './sheet';
window.onload = function() {
    let spread = new GC.Spread.Sheets.Workbook(_getElementById('ss'));
    let dataSource1 = dataSource;
    spread.setSheetCount(dataSource1.length);
    let spreadNS = GC.Spread.Sheets;
    spread.suspendPaint();
    spread.suspendCalcService(false);
    new Sheet(3, spreadNS, spread, dataSource1);
    new Sheet(1, spreadNS, spread, dataSource1);
    new Sheet(2, spreadNS, spread, dataSource1);
    new Sheet(0, spreadNS, spread, dataSource1);
    spread.resumeCalcService(false);
    spread.resumePaint();

};

function _getElementById(id) {
    return document.getElementById(id);
}