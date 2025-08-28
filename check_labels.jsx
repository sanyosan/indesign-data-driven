function readCSVRows(file) {
  var rows = [];
  var buffer = "";
  while (!file.eof) {
    var line = file.readln();
    buffer += (buffer === "" ? "" : "\n") + line;
    var quoteCount = (buffer.match(/"/g) || []).length;
    if (quoteCount % 2 === 0) {
      rows.push(buffer);
      buffer = "";
    }
  }
  if (buffer !== "") rows.push(buffer);
  return rows;
}

function parseCSVLine(line) {
  var result = [];
  var inQuotes = false;
  var value = "";
  for (var i = 0; i < line.length; i++) {
    var ch = line[i];
    if (ch === '"') {
      inQuotes = !inQuotes;
    } else if (ch === ',' && !inQuotes) {
      result.push(value);
      value = "";
    } else {
      value += ch;
    }
  }
  result.push(value);
  for (var j = 0; j < result.length; j++) {
    var v = result[j];
    if (v.length > 1 && v[0] === '"' && v[v.length-1] === '"') {
      v = v.substring(1, v.length-1).replace(/""/g, '"');
    }
    // 改行コードをInDesign用に変換
    v = v.replace(/\n/g, "\r");
    result[j] = v;
  }
  return result;
}

// 位置設定CSVを読み込む関数
function readPositionCSV(scriptPath) {
  var posFile = File(scriptPath + "/pos.csv");
  if (!posFile.exists) {
    alert("pos.csvが見つかりません: " + posFile.fsName);
    return null;
  }
  
  posFile.open("r");
  var posRows = readCSVRows(posFile);
  posFile.close();
  
  var positions = [];
  for (var i = 1; i < posRows.length; i++) { // ヘッダーをスキップ
    var cols = parseCSVLine(posRows[i]);
    if (cols.length >= 3) {
      positions.push({
        name: cols[0],
        x: parseFloat(cols[1]) || 0,
        y: parseFloat(cols[2]) || 0
      });
    }
  }
  return positions;
}

function myTrim(str) { 
  return (str || "").replace(/^\s+|\s+$/g, ""); 
}

#target "indesign"

(function () {
  if (app.documents.length === 0) { 
    alert("ドキュメントを開いてください"); 
    return; 
  }
  var doc = app.activeDocument;

  if (doc.selection.length !== 1 || !(doc.selection[0] instanceof Group)) {
    alert("グループを1つ選択してください。");
    return;
  }

  var baseGroup = doc.selection[0];
  var gb = baseGroup.geometricBounds;
  var baseX = gb[1];
  var baseY = gb[0];

  // スクリプトのパスを取得
  var scriptPath = File($.fileName).parent.toString();
  
  // 位置設定CSVを読み込み
  var positions = readPositionCSV(scriptPath);
  if (!positions) return;

  // データCSVを選択して読み込み
  var dataFile = File.openDialog("データCSVファイルを選択してください", "*.csv");
  if (!dataFile) return;
  
  dataFile.open("r");
  var dataRows = [];
  var csvLines = readCSVRows(dataFile);
  for (var i = 0; i < csvLines.length; i++) {
    var cols = parseCSVLine(csvLines[i]);
    if (cols.length > 1 || (cols.length === 1 && cols[0] !== "")) {
      dataRows.push(cols);
    }
  }
  dataFile.close();

  if (dataRows.length < 2) { 
    alert("データ行がありません"); 
    return; 
  }

  var header = dataRows[0];
  var dataValues = dataRows.slice(1);
  
  // 処理できる最大数を計算
  var maxItems = Math.min(positions.length, dataValues.length);
  
  alert("処理を開始します。" + maxItems + "個のグループを作成し、pos.csvの位置に配置します。");

  // 各データ行に対してグループを作成・配置
  for (var idx = 0; idx < maxItems; idx++) {
    var row = dataValues[idx];
    var pos = positions[idx];
    
    // 最初の要素は元のグループを使用、以降は複製
    var newGroup = (idx === 0) ? baseGroup : baseGroup.duplicate();

    // pos.csvの座標に基づいて配置
    var targetX = baseX + pos.x;
    var targetY = baseY + pos.y;
    
    // 現在の位置を取得して移動量を計算
    var currentBounds = newGroup.geometricBounds;
    var deltaX = targetX - currentBounds[1];
    var deltaY = targetY - currentBounds[0];
    newGroup.move([deltaX, deltaY]);

    // グループ内のTextFrameを検索してテキストを置換
    var items = newGroup.allPageItems;
    for (var i = 0; i < items.length; i++) {
      if (!(items[i] instanceof TextFrame)) continue;
      var label = myTrim(items[i].label);
      
      // ラベルが空の場合はスキップ
      if (!label) continue;

      // ヘッダーでラベルに一致する列を検索
      var colIndex = -1;
      for (var h = 0; h < header.length; h++) {
        if (myTrim(header[h]) === label) {
          colIndex = h;
          break;
        }
      }
      
      // 一致する列が見つかった場合、データを設定
      if (colIndex >= 0) {
        items[i].contents = row[colIndex] || "";
      }
    }
    
    // 進捗表示
    if (idx % 3 === 2) {
      app.activeDocument.recompose();
    }
  }

  app.activeDocument.recompose();
  alert("データ流し込み完了！" + maxItems + "個のグループを作成し、pos.csvの位置に配置しました。");
})();