// InDesign ExtendScript: 8つのオブジェクトを順番に選択して相対座標をCSV出力
// data1からdata8までの配置位置を記録

var positions = [];
var baseX = 0;
var baseY = 0;

// InDesign ExtendScript: 8つのオブジェクトを一度に選択して相対座標をCSV出力
// data1からdata8までの配置位置を記録

var positions = [];
var baseX = 0;
var baseY = 0;

// 現在選択されているオブジェクトを取得
var sel = app.activeDocument.selection;

if (sel.length < 8) {
    alert("8つのオブジェクトを選択してください。\n現在選択数: " + sel.length);
} else {
    alert("8つのオブジェクトが選択されました。\n選択順にdata1〜data8として処理します。");
    
    // 選択されたオブジェクトを順番に処理
    for (var i = 0; i < 8; i++) {
        var bounds = sel[i].visibleBounds; // [y1, x1, y2, x2]
        var x = bounds[1]; // 左端
        var y = bounds[0]; // 上端
        
        if (i === 0) {
            // 最初の位置を基準とする
            baseX = x;
            baseY = y;
            positions.push({name: "data" + (i + 1), x: 0, y: 0});
        } else {
            // 相対位置を計算
            var relX = x - baseX;
            var relY = y - baseY;
            positions.push({name: "data" + (i + 1), x: relX, y: relY});
        }
    }

    // CSVを生成
    var csv = "name,x,y\n";
    for (var i = 0; i < positions.length; i++) {
        csv += positions[i].name + "," + positions[i].x + "," + positions[i].y + "\n";
    }

    // ファイル保存
    var file = File.saveDialog("位置情報CSVを保存", "*.csv");
    if (file) {
        file.encoding = "UTF-8";
        file.open("w");
        file.write(csv);
        file.close();
        alert("位置情報CSV保存完了\n8つの位置が記録されました");
    }
}
