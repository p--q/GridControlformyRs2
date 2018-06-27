#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from indoc import historydialog1, historydialog2, historydialog3, historydialog4, historydialog5, historydialog6, historydialog7, staticdialog1
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
						if dialogname=="historydialog1":  # 履歴ダイアログ。ダイアログを閉じる時に重複要素を削除する。
							historydialog1.createDialog(xscriptcontext, enhancedmouseevent, dialogname, defaultrows)	
						elif dialogname=="historydialog2":   # 履歴ダイアログ。グリッドデータを変更する時に重複を削除する。他バグフィックス。
							historydialog2.createDialog(xscriptcontext, enhancedmouseevent, dialogname, defaultrows)	
						elif dialogname=="historydialog3":   # 履歴ダイアログ。逐次検索機能追加。
							historydialog3.createDialog(xscriptcontext, enhancedmouseevent, dialogname, defaultrows)	
						elif dialogname=="historydialog4":   # 履歴ダイアログ。リファクタリング。削除のバグフィクス。
							historydialog4.createDialog(xscriptcontext, enhancedmouseevent, dialogname, defaultrows)						
						elif dialogname=="staticdialog1":   # 静的ダイアログ。
							staticdialog1.createDialog(xscriptcontext, enhancedmouseevent, dialogname, defaultrows)					
						elif dialogname=="historydialog5":   # 履歴ダイアログ。余白の縮小。逐語検索したあと1行だけになると、その行をマウスで選択しないと、ダイアログがマウス操作を受け付けない。
							# 逐語検索を有効にして2行目以降の行を選択して１行だけになった時に、その行をマウスで選択しないとダイアログ上の他の操作ができなくなる。
							historydialog5.createDialog(xscriptcontext, enhancedmouseevent, dialogname, defaultrows)					
						elif dialogname=="historydialog6":   # 履歴ダイアログ。没。インデックスで絞り込み使用としたが難しいので断念。2段階目のインデックス取得が難しい。
							historydialog6.createDialog(xscriptcontext, enhancedmouseevent, dialogname, defaultrows)					
						elif dialogname=="historydialog7":   # 履歴ダイアログ。GridSelectionListenerの使用中止。
							historydialog7.createDialog(xscriptcontext, enhancedmouseevent, dialogname, defaultrows)					
					
						
					return False  # セル編集モードにしない。
		return True  # セル編集モードにする。
