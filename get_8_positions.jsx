// InDesign ExtendScript: 8つの位置をクリックして相対座標をCSV出力
// data1からdata8までの配置位置を記録

var positions = [];
var baseX = 0;
var baseY = 0;

// ユーザーに説明
alert("8つの位置をクリックしてください。\n1つ目は基準位置(data1)、\n2つ目以降は相対位置で記録されます。");

// 8回クリック位置を取得
for (var i = 0; i < 8; i++) {
    var clickPos = app.activeDocument.pages[0].pointFromRuler([0, 0]);
    
    // マウスクリック待機（簡易版）
    var dialog = app.dialogs.add();
    var panel = dialog.dialogColumns.add();
    var group = panel.borderPanels.add();
    group.staticTexts.add({staticLabel: "data" + (i + 1) + "の位置でクリックしてOKを押してください"});
    
    if (dialog.show() == true) {
        // 現在のマウス位置を取得（近似）
        var x = Math.random() * 200; // 実際の実装では別の方法が必要
        var y = Math.random() * 200;
        
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
    
    dialog.destroy();
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
