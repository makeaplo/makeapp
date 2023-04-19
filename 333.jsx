

// 获取关键帧数据
function getKeyframeData(layer, effectName) {
    var keyframeData = [];

    // 遍历图层属性
    function processProperties(group, prefix, effectIndex) {
        for (var propIndex = 1; propIndex <= group.numProperties; propIndex++) {
            var property = group.property(propIndex);

            if (property.propertyType === PropertyType.PROPERTY && property.numKeys > 0) {
                var propertyName = prefix + property.name;
                for (var keyIndex = 1; keyIndex <= property.numKeys; keyIndex++) {
                    var keyTime = property.keyTime(keyIndex);
                    var keyValue = property.keyValue(keyIndex);
                    keyframeData.push({
                        propertyName: propertyName,
                        keyIndex: keyIndex,
                        keyTime: keyTime,
                        keyValue: keyValue
                    });
                }
            } else if (property.propertyType === PropertyType.INDEXED_GROUP || property.propertyType === PropertyType.NAMED_GROUP) {
            if (property.name.includes(effectName)) {
                processProperties(property, prefix + property.name + ' ' + effectIndex + ' - ', effectIndex);
            }

            }
        }
    }

    processProperties(layer, '', 1);
    return keyframeData;
}



// 将关键帧数据导出为CSV文件
function exportKeyframeDataToCSV(data, compFrameRate, filePath) {
    // 获取所有不同的属性名称
    var propertyNames = data.map(function (d) {
        return d.propertyName;
    }).filter(function (value, index, self) {
        return self.indexOf(value) === index;
    });

    // 生成CSV文件的第一行（标题行）
    var csvContent = 'Key Index,Key Frame,';
    for (var i = 1; i <= propertyNames.length; i++) {
        csvContent += i + ',';
    }
    csvContent += '\n';


    // 根据关键帧时间对数据进行分组
    var keyframeDataByTime = {};
    for (var i = 0; i < data.length; i++) {
        var keyTime = data[i].keyTime;
        var keyFrame = Math.round(keyTime * compFrameRate); // 转换为帧数
        if (!keyframeDataByTime.hasOwnProperty(keyFrame)) {
            keyframeDataByTime[keyFrame] = {};
        }
        keyframeDataByTime[keyFrame][data[i].propertyName] = data[i].keyValue;
    }

    // 生成CSV文件的其他行（数据行）
    for (var keyFrame in keyframeDataByTime) {
        var row = [keyFrame];
        for (var j = 0; j < propertyNames.length; j++) {
            var propertyName = propertyNames[j];
            if (keyframeDataByTime[keyFrame].hasOwnProperty(propertyName)) {
                row.push(keyframeDataByTime[keyFrame][propertyName]);
            } else {
                row.push('');
            }
        }
        csvContent += row.join(',') + '\n';
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
    var fileName = 'keyframe_data.csv';
    var filePath = desktopPath + '/' + fileName;

    var compFrameRate = activeItem.frameRate;
    exportKeyframeDataToCSV(keyframeData, compFrameRate, filePath);

    alert('Keyframe data has been exported to: ' + filePath);
}

main();
