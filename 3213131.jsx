// 获取所有合成
var allComps = [];
for (var i = 1; i <= app.project.numItems; i++) {
  if (app.project.item(i) instanceof CompItem) {
    allComps.push(app.project.item(i));
  }
}

// 获取每个合成中的 MusicData 图层的关键帧数据
var allData = [];
for (var i = 0; i < allComps.length; i++) {
  var currentComp = allComps[i];
  var currentData = getDataFromComp(currentComp);
  if (currentData.length > 0) {
    allData.push(currentData);
  }
}

// 将所有数据导出到 Excel 文件中
exportToExcel(allData);

// 从合成中获取 MusicData 图层的关键帧数据
function getDataFromComp(comp) {
  var musicDataLayer = null;
  for (var i = 1; i <= comp.numLayers; i++) {
    if (comp.layer(i) instanceof AVLayer && comp.layer(i).name == "MusicData") {
      musicDataLayer = comp.layer(i);
      break;
    }
  }
  if (!musicDataLayer) {
    return [];
  }

  var data = {
    compName: comp.name,
    layerName: musicDataLayer.name,
    properties: [],
  };

  for (var i = 1; i <= musicDataLayer.numProperties; i++) {
    var prop = musicDataLayer.property(i);
    if (prop.numKeys > 0) {
      var propData = getPropertyData(prop);
      data.properties.push(propData);
    }
  }

  return data;
}

// 获取属性的关键帧数据
function getPropertyData(prop) {
  var propData = {
    propName: prop.name,
    keyframes: [],
  };

  for (var i = 1; i <= prop.numKeys; i++) {
    var keyData = {
      time: prop.keyTime(i),
      value: prop.keyValue(i),
      easeIn: prop.keyInTemporalEase(i),
      easeOut: prop.keyOutTemporalEase(i),
    };
    propData.keyframes.push(keyData);
  }

  return propData;
}

// 导出数据到 Excel 文件
function exportToExcel(allData) {
  // 创建 Excel 对象
  var excel = new ActiveXObject("Excel.Application");
  excel.visible = true;
  var workbook = excel.Workbooks.Add();
  var sheet = workbook.Sheets.Add();

  // 添加表头
  var headers = [
    "Comp Name",
    "Layer Name",
    "Property Name",
    "Time",
    "Value",
    "Ease In",
    "Ease Out",
  ];
  for (var i = 0; i < headers.length; i++) {
    sheet.Cells(1, i + 1).Value = headers[i];
  }

  // 添加数据
  var row = 2;
  for (var i = 0; i < allData.length; i++) {
    var currentData = allData[i];
    var compName = currentData.compName;
    var layerName = currentData.layerName;
    for (var j = 0; j < currentData.properties.length; j++) {
      var currentProp = currentData.properties[j];
      var propName = currentProp.propName;
      for (var k = 0;
var excelFile = new File("~/Desktop/MusicData.xlsx");
var workbook = XLSX.readFile(excelFile.fsName);
var sheetName = "MusicData";

function getMusicDataLayers(composition) {
  var layers = composition.layers;
  var musicDataLayers = [];

  for (var i = 1; i <= layers.length; i++) {
    var layer = layers[i];
    if (layer.name === "MusicData" && layer.enabled) {
      musicDataLayers.push(layer);
    } else if (layer instanceof AVLayer) {
      continue;
    } else if (layer instanceof AVLayer) {
      musicDataLayers = musicDataLayers.concat(getMusicDataLayers(layer.source));
    }
  }

  return musicDataLayers;
}

function exportMusicData() {
  var compositions = app.project.items;
  var musicData = [];

  for (var i = 1; i <= compositions.length; i++) {
    var composition = compositions[i];
    if (composition instanceof CompItem) {
      var musicDataLayers = getMusicDataLayers(composition);
      if (musicDataLayers.length > 0) {
        var sheet = workbook.Sheets[sheetName] || workbook.Sheets.Add(sheetName);
        var rowData = [];

        for (var j = 0; j < musicDataLayers.length; j++) {
          var layer = musicDataLayers[j];

          var group = layer.property("ADBE Effect Parade");
          if (group) {
            group.forEach(function (effect, index) {
              var name = effect.name;
              var props = effect.properties;
              props.forEach(function (prop) {
                if (prop.numKeys > 0) {
                  var keyTimes = [];
                  var keyValues = [];
                  for (var k = 1; k <= prop.numKeys; k++) {
                    keyTimes.push(prop.keyTime(k));
                    keyValues.push(prop.keyValue(k));
                  }
                  rowData.push({
                    Composition: composition.name,
                    Layer: layer.name,
                    Effect: name,
                    Property: prop.name,
                    KeyTimes: keyTimes,
                    KeyValues: keyValues,
                  });
                }
              });
            });
          }
        }

        for (var j = 0; j < rowData.length; j++) {
          var row = rowData[j];
          var rowIndex = sheet["!rows"] ? sheet["!rows"].length : 0;
          XLSX.utils.sheet_add_json(sheet, [row], { skipHeader: rowIndex > 0, origin: rowIndex });
        }
      }
    }
  }

  XLSX.writeFile(workbook, excelFile.fsName);
}

exportMusicData();
// 将数据写入到excel中
function writeToExcel(sheetName, data) {
  var desktop = Folder.desktop;
  var excelFile = new File(desktop.fsName + "/" + sheetName + ".xlsx");

  // 新建excel文件
  var excel = new ActiveXObject("Excel.Application");
  var excelBook = excel.Workbooks.Add();
  var excelSheet = excelBook.Worksheets(1);
  excelSheet.Name = sheetName;

  // 写入数据
  for (var i = 0; i < data.length; i++) {
    var rowData = data[i];
    for (var j = 0; j < rowData.length; j++) {
      excelSheet.Cells(i + 1, j + 1).value = rowData[j];
    }
  }

  // 保存并关闭excel
  excelBook.SaveAs(excelFile);
  excelBook.Close();
  excel.Quit();
}

// 主程序
var selectedComps = getSelectedComps();
for (var i = 0; i < selectedComps.length; i++) {
  var comp = selectedComps[i];
  var musicDataLayers = getMusicDataLayers(comp);
  if (musicDataLayers.length > 0) {
    var sheetName = comp.name;
    var data = getDataFromMusicDataLayers(musicDataLayers);
    writeToExcel(sheetName, data);
  }
}

alert("导出完成！");
