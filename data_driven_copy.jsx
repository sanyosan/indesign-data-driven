// InDesign ExtendScript: 選択グループをpos.csvの位置にコピーし、data.csvの内容でテキスト置換

// CSVファイルを読み込む関数
function readCSV(filePath) {
    var file = File(filePath);
    if (!file.exists) {
        alert("ファイルが見つかりません: " + filePath);
        return null;
    }
    
    file.open("r");
    var content = file.read();
    file.close();
    
    var lines = content.split("\n");
    var result = [];
    var headers = lines[0].split(",");
    
    for (var i = 1; i < lines.length; i++) {
        if (lines[i] === "" || lines[i] === "\r" || lines[i] === "\n") continue;
        var values = lines[i].split(",");
        var row = {};
        for (var j = 0; j < headers.length; j++) {
            row[headers[j]] = values[j] || "";
        }
        result.push(row);
    }
    
    return result;
}

// グループ内のTextFrameを再帰的に検索して置換する関数
function replaceTextInGroup(group, dataRow) {
    for (var i = 0; i < group.pageItems.length; i++) {
        var item = group.pageItems[i];
        if (item.hasOwnProperty('contents')) {
            // TextFrameの場合
            var label = item.label;
            // スクリプトラベルをデバッグ表示
            if (label && label !== "") {
                // ラベルがdata.csvのヘッダーと一致する場合に置換
                if (dataRow[label]) {
                    item.contents = dataRow[label];
                }
            }
        } else if (item.constructor.name === "Group") {
            // 再帰的にグループ内を検索
            replaceTextInGroup(item, dataRow);
        }
    }
}

// メイン処理開始
var sel = app.activeDocument.selection[0];
if (!sel || sel.constructor.name !== "Group") {
    alert("グループを選択してください");
    exit();
}

var baseGroup = sel;
var baseBounds = baseGroup.visibleBounds;
var baseX = baseBounds[1];
var baseY = baseBounds[0];

// CSVファイル読み込み
var scriptPath = File($.fileName).parent.toString();
var posData = readCSV(scriptPath + "/pos.csv");
var dataData = readCSV(scriptPath + "/data.csv");

if (!posData || !dataData) {
    alert("CSVファイルの読み込みに失敗しました");
    exit();
}

alert("処理を開始します。" + posData.length + "個の位置にグループをコピーし、data.csvの行を順番に流し込みます。");

// pos.csvの順序でグループをコピーし、data.csvの行を順番に適用
for (var i = 0; i < Math.min(posData.length, dataData.length); i++) {
    var pos = posData[i];
    var data = dataData[i]; // data.csvの行を順番に使用（行1→data5位置、行2→data6位置...）
    
    // グループをコピー&ペースト方式で複製（スクリプトラベル保持）
    baseGroup.select();
    app.copy();
    app.paste();
    var newGroup = app.activeDocument.selection[0];
    
    // 新しい位置を計算（選択したグループの左上 + 相対位置）
    var newX = baseX + parseFloat(pos.x);
    var newY = baseY + parseFloat(pos.y);
    
    // グループを新しい位置に配置
    var currentBounds = newGroup.visibleBounds;
    var currentX = currentBounds[1]; // 現在の左端
    var currentY = currentBounds[0]; // 現在の上端
    
    // 移動量を計算
    var deltaX = newX - currentX;
    var deltaY = newY - currentY;
    newGroup.move([deltaX, deltaY]);
    
    // グループ内のTextFrameのテキストを置換
    replaceTextInGroup(newGroup, data);
    
    // 進捗表示（どの位置にどの行のデータを適用したかを表示）
    alert("位置: " + pos.name + " にdata.csv行" + (i + 1) + "のデータを適用しました");
    
    // 進捗表示
    if (i % 2 === 1) {
        app.activeDocument.recompose();
    }
}

app.activeDocument.recompose();
alert("処理完了！" + Math.min(posData.length, dataData.length) + "個のグループを作成し、テキストを置換しました。");
