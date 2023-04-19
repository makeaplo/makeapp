// 获取关键帧数据
function getKeyframeData(layer) {
    var keyframeData = [];

    // 遍历图层属性
    function processProperties(group) {
        for (var propIndex = 1; propIndex <= group.numProperties; propIndex++) {
            var property = group.property(propIndex);

            if (property.propertyType === PropertyType.PROPERTY && property.numKeys > 0) {
                for (var keyIndex = 1; keyIndex <= property.numKeys; keyIndex++) {
                    //var keyTime = property.keyTime(keyIndex);
                    var keyValue = property.keyValue(keyIndex);
                    keyframeData.push({
                        //propertyName: property.name,
                        keyIndex: keyIndex,
                        //keyTime: keyTime,
                        keyValue: keyValue
                    });
                }
            } else if (property.propertyType === PropertyType.INDEXED_GROUP || property.propertyType === PropertyType.NAMED_GROUP) {
                processProperties(property);
            }
        }
    }

    processProperties(layer);
    return keyframeData;
}


// 将关键帧数据导出为CSV文件
function exportKeyframeDataToCSV(data, filePath) {
    var csvContent = 'Key Index,Key Value\n';

    for (var i = 0; i < data.length; i++) {
        var row = [
            //data[i].propertyName,
            data[i].keyIndex,
            //data[i].keyTime,
            data[i].keyValue
        ].join(',');
        csvContent += row + '\n';
    }

    var file = new File(filePath);
    file.open('w');
    file.write(csvContent);
    file.close();
}

// 主程序
function main() {
    var activeItem = app.project.activeItem;
    if (activeItem == null || !(activeItem instanceof CompItem)) {
        alert('Please select a composition with a layer to export keyframes.');
        return;
    }

    var selectedLayers = activeItem.selectedLayers;
    if (selectedLayers.length === 0) {
        alert('Please select a layer to export keyframes.');
        return;
    }

    var selectedLayer = selectedLayers[0];
    var keyframeData = getKeyframeData(selectedLayer);

    if (keyframeData.length === 0) {
        alert('No keyframes found on the selected layer.');
        return;
    }

    var desktopPath = Folder.desktop.fullName;
    var fileName = 'keyframe_data1_24.csv';
    var filePath = desktopPath + '/' + fileName;

    exportKeyframeDataToCSV(keyframeData, filePath);
    alert('Keyframe data has been exported to: ' + filePath);
}

main();
