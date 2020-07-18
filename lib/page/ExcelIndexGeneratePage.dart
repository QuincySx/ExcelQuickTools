import 'dart:io';

import 'package:excel/excel.dart';
import 'package:filepicker_windows/filepicker_windows.dart';
import 'package:flutter/cupertino.dart';
import 'package:flutter/material.dart';

class ExcelIndexGeneratePage extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text("生成填写 Excel 序号"),
      ),
      body: ExcelPage(),
    );
  }
}

class ExcelPage extends StatefulWidget {
  @override
  State<StatefulWidget> createState() {
    return _ExcelPageState();
  }
}

class _ExcelPageState extends State<ExcelPage> {
  File path;
  String mCellLocation;
  String mContentPrefix;
  String mContentSuffix;
  var isShowProgress = false;

  @override
  void initState() {
    super.initState();
  }

  void _processFile() {
    if (path == null) {
      _showAlertDialog("请先选择要处理的 Excel 文件");
      return;
    }

    if (mCellLocation == null) {
      _showAlertDialog("请输入要批量输入的单元格");
      return;
    }

    if (mContentSuffix == null) {
      _showAlertDialog("请输入起始编号");
      return;
    }

    var number = int.tryParse(mContentSuffix);
    if (number == null) {
      _showAlertDialog("起始编号请输入数字");
      return;
    }

    var digitalRegExp = new RegExp(r"[0-9]");
    var row = digitalRegExp.allMatches(mCellLocation).map((m) => m.group(0));

    var letterRegExp = new RegExp(r"[a-zA-Z]");
    var cell = letterRegExp.allMatches(mCellLocation).map((m) => m.group(0));

    if (row == null || cell == null || row.length == 0 || cell.length == 0) {
      _showAlertDialog("输入的单元格位置有误");
      return;
    }

    setState(() {
      isShowProgress = true;
    });
    var numberCount = mContentSuffix.length;
    var bytes = File(path.path).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);

    for (var i = 0; i < excel.tables.values.length; i++) {
      var sheet = excel.tables.values.elementAt(i);
      var sheetNumber = number + i;
      var cell = sheet.cell(CellIndex.indexByString(mCellLocation));
      cell.value = (mContentPrefix == null ? "" : mContentPrefix) +
          sheetNumber.toString().padLeft(numberCount, '0');
    }

    var dotIndex = path.path.lastIndexOf('.');
    var pathPrefix = path.path.substring(0, dotIndex) + "-new";
    var pathSuffix = path.path.substring(dotIndex);
    var newPath = pathPrefix + pathSuffix;
    excel.encode().then((onValue) {
      File(newPath)
        ..createSync(recursive: true)
        ..writeAsBytesSync(onValue);
    });

    setState(() {
      isShowProgress = false;
    });

    _showAlertDialog("Excel 文件已经保存，新的路径为 \n$newPath");
  }

  void _showAlertDialog(String msg) {
    showDialog(
        context: context,
        barrierDismissible: false,
        builder: (BuildContext context) {
          return AlertDialog(
            title: Text('提示'),
            //可滑动
            content: SingleChildScrollView(
              child: Text(msg),
            ),
            actions: <Widget>[
              FlatButton(
                child: Text('确定'),
                onPressed: () {
                  Navigator.of(context).pop();
                },
              ),
            ],
          );
        });
  }

  @override
  Widget build(BuildContext context) {
    if (isShowProgress) {
      return Center(
        child: CupertinoActivityIndicator(
          radius: 30.0, //值越大加载的图形越大
        ),
      );
    }
    return Center(
      child: Column(
        mainAxisAlignment: MainAxisAlignment.center,
        children: [
          Column(
            children: [
              Text(
                "选中的 Excel 文件",
                style: new TextStyle(
                  inherit: true,
                  color: Color(0xFF555555),
                  fontSize: 22.0,
                ),
              ),
              Padding(
                padding: EdgeInsets.only(
                  top: 8.0,
                  bottom: 24.0,
                ),
                child: Text(path == null ? "未选择" : path.path),
              ),
            ],
          ),
          CupertinoButton(
            color: Color(0xFF1AAD19),
            child: Text("选择 Excel 文件"),
            onPressed: () {
              final file = FilePicker();
              file.hidePinnedPlaces = true;
              file.forcePreviewPaneOn = true;
              file.filterSpecification = {
                'Excel Files': '*.xls;*.xlsx',
              };
              file.title = 'Select an image';
              final result = file.getFile();
              if (result != null) {
                setState(() {
                  path = result;
                });
              }
            },
          ),
          Container(
            width: 140.0,
            child: TextField(
              textAlign: TextAlign.center,
              style: TextStyle(
                fontSize: 18,
              ),
              decoration: InputDecoration(
                labelText: "单元格位置",
                hintText: "例子：B3",
              ),
              onChanged: (value) {
                setState(() {
                  mCellLocation = value;
                });
              },
            ),
          ),
          Container(
            width: 140.0,
            child: TextField(
              textAlign: TextAlign.center,
              style: TextStyle(
                fontSize: 18,
              ),
              decoration: InputDecoration(
                labelText: "前缀",
                hintText: "例子：Number-",
              ),
              onChanged: (value) {
                setState(() {
                  mContentPrefix = value;
                });
              },
            ),
          ),
          Container(
            width: 140.0,
            child: TextField(
              textAlign: TextAlign.center,
              style: TextStyle(
                fontSize: 18,
              ),
              decoration: InputDecoration(
                labelText: "起始编号",
                hintText: "例子：01",
              ),
              onChanged: (value) {
                setState(() {
                  mContentSuffix = value;
                });
              },
            ),
          ),
          Padding(
            padding: EdgeInsets.only(
              top: 24.0,
            ),
            child: Row(
              mainAxisAlignment: MainAxisAlignment.center,
              children: [
                Text("示例："),
                Text((mContentPrefix == null ? '' : mContentPrefix) +
                    (mContentSuffix == null ? '' : mContentSuffix)),
              ],
            ),
          ),
          Padding(
            padding: EdgeInsets.only(
              top: 24.0,
            ),
            child: CupertinoButton(
              color: Color(0xFF4169E1),
              child: Text("将编号写入 Excel"),
              onPressed: () {
                _processFile();
              },
            ),
          ),
        ],
      ),
    );
  }
}
