import 'package:flutter/cupertino.dart';
import 'package:flutter/material.dart';

class MyHomePage extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: Text("Excel 快捷工具")),
      body: Center(
        child: Column(
          children: [
            Padding(
              padding: EdgeInsets.all(16.0),
            ),
            CupertinoButton(
              color: Color(0xFF4169E1),
              child: Text("Excel 添加序号索引"),
              onPressed: () {
                Navigator.pushNamed(context, '/ExcelIndexGeneratePage');
              },
            ),
          ],
        ),
      ),
    );
  }
}
