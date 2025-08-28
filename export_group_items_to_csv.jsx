// InDesign ExtendScript: グループ内のTextFrame/Rectangleの相対座標・内容をCSV出力
// グループを選択して実行してください

var sel = app.activeDocument.selection[0];
if (sel.constructor.name !== "Group") {
    alert("グループを選択してください");
    exit();
}
var group = sel;
var groupBounds = group.visibleBounds; // [y1, x1, y2, x2]
var groupX = groupBounds[1];
var groupY = groupBounds[0];

var sel = app.activeDocument.selection[0];
if (sel.constructor.name !== "Group") {
    alert("グループを選択してください");
    exit();
}
var group = sel;
var groupBounds = group.visibleBounds; // [y1, x1, y2, x2]
var groupX = groupBounds[1];
var groupY = groupBounds[0];

var csv = "type,x,y,width,height,text,imageFile\n";

for (var i = 0; i < group.pageItems.length; i++) {
    var item = group.pageItems[i];
    var bounds = item.visibleBounds;
    var x = bounds[1] - groupX;
    var y = bounds[0] - groupY;
    var w = bounds[3] - bounds[1];
    var h = bounds[2] - bounds[0];

    var type = item.constructor.name;
    var text = "";
    var imageFile = "";

    // より正確なタイプ判定
    if (item.hasOwnProperty('contents')) {
        type = "TextFrame";
    } else if (item.hasOwnProperty('images') && item.images.length > 0) {
        type = "Image";
    } else if (item.hasOwnProperty('rectangles') || item.hasOwnProperty('ovals') || item.hasOwnProperty('polygons')) {
        type = "Shape";
    }

    // デバッグ：すべてのオブジェクトタイプを出力
    if (type === "TextFrame" || item.hasOwnProperty('contents')) {
        // 複数の方法でテキストを取得を試行
        try {
            if (item.contents) {
                text = item.contents.replace(/(\r|\n)/g, " ");
            } else if (item.parentStory && item.parentStory.contents) {
                text = item.parentStory.contents.replace(/(\r|\n)/g, " ");
            } else if (item.texts.length > 0) {
                text = item.texts[0].contents.replace(/(\r|\n)/g, " ");
            } else {
                text = "テキスト取得失敗";
            }
        } catch (e) {
            text = "エラー:" + e.message;
        }
        csv += type + "," + x + "," + y + "," + w + "," + h + "," + text + "," + imageFile + "\n";
    } else if (type === "Rectangle") {
        if (item.images.length > 0) {
            if (item.images[0].itemLink) {
                imageFile = item.images[0].itemLink.name;
            } else {
                imageFile = "埋め込み画像またはリンクなし";
            }
        }
        csv += type + "," + x + "," + y + "," + w + "," + h + "," + text + "," + imageFile + "\n";
    } else {
        // その他のオブジェクトも表示
        csv += type + "," + x + "," + y + "," + w + "," + h + "," + text + "," + imageFile + "\n";
    }
}

var file = File.saveDialog("CSVを保存", "*.csv");
if (file) {
    file.encoding = "UTF-8";
    file.open("w");
    file.write(csv);
    file.close();
    alert("CSV保存完了");
}
