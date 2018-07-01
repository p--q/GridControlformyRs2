#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from indoc import staticdialog3, historydialog8, datedialog1
from com.sun.star.awt import MouseButton  # 定数
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
		selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
			if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
				if enhancedmouseevent.ClickCount==2:  # ダブルクリックの時。
					sheet = selection.getSpreadsheet()
					celladdress = selection.getCellAddress()
					r, c = celladdress.Row, celladdress.Column
					dialogname = sheet[0, c].getString()
					if r>0:
						defaultrows = "item1", "item2", "item3", "item4"			
						if dialogname=="staticdialog3":   # 静的ダイアログ。ポップアップメニューアイテムを名前で取得に変更。
							staticdialog3.createDialog(xscriptcontext, enhancedmouseevent, dialogname, defaultrows)				
						elif dialogname=="historydialog8":   # 履歴ダイアログ。選択行インデックスの取得方法、スクロール、を修正。
							historydialog8.createDialog(xscriptcontext, enhancedmouseevent, dialogname, defaultrows)		
						elif dialogname=="datedialog1":  # 日付ダイアログ。
							datedialog1.createDialog(xscriptcontext, enhancedmouseevent, dialogname)		

					
					return False  # セル編集モードにしない。
		return True  # セル編集モードにする。
