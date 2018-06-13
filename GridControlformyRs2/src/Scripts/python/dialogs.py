#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from datetime import datetime
import json
from com.sun.star.accessibility import AccessibleRole  # 定数
from com.sun.star.awt import XActionListener
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.awt import XMenuListener
from com.sun.star.awt import XMouseListener
from com.sun.star.awt import MouseButton  # 定数
from com.sun.star.awt import PopupMenuDirection  # 定数
from com.sun.star.awt import Rectangle  # Struct
from com.sun.star.awt import ScrollBarOrientation  # 定数
from com.sun.star.document import XDocumentEventListener
from com.sun.star.util import XCloseListener
from com.sun.star.view.SelectionType import MULTI  # enum 
from com.sun.star.awt.MessageBoxType import QUERYBOX  # enum
from com.sun.star.awt import MessageBoxButtons  # 定数
from com.sun.star.awt import MessageBoxResults  # 定数
from com.sun.star.awt.grid import XGridSelectionListener
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.awt import Point  # Struct
from com.sun.star.util import MeasureUnit  # 定数
from com.sun.star.frame import XFrameActionListener
from com.sun.star.frame.FrameAction import FRAME_UI_DEACTIVATING  # enum
def macro(documentevent=None):  # 引数は文書のイベント駆動用。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
	controller = doc.getCurrentController()  # コントローラの取得。
	enhancedmouseclickhandler = EnhancedMouseClickHandler(XSCRIPTCONTEXT, controller)
	controller.addEnhancedMouseClickHandler(enhancedmouseclickhandler)  # EnhancedMouseClickHandler	
	doc.addDocumentEventListener(DocumentEventListener(enhancedmouseclickhandler))  # DocumentEventListener。ドキュメントのリスナーの削除のため。	
class DocumentEventListener(unohelper.Base, XDocumentEventListener):
	def __init__(self, enhancedmouseclickhandler):
		self.args = enhancedmouseclickhandler
	def documentEventOccured(self, documentevent):  # ドキュメントのリスナーを削除する。
		enhancedmouseclickhandler = self.args
		if documentevent.EventName=="OnUnload":  
			source = documentevent.Source
			source.removeEnhancedMouseClickHandler(enhancedmouseclickhandler)
			source.removeDocumentEventListener(self)
	def disposing(self, eventobject):
		eventobject.Source.removeDocumentEventListener(self)
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler):
	def __init__(self, xscriptcontext, subj):
		self.subj = subj
		self.xscriptcontext = xscriptcontext
	def mousePressed(self, enhancedmouseevent):  # try文を使わないと1回のエラーで以後動かなくなる。
		xscriptcontext = self.xscriptcontext
		selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
			if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
				if enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
					try:
# 						txt =selection.getString()
						createDialog(xscriptcontext, enhancedmouseevent)		
						return False  # セル編集モードにしない。
					except:
						import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
		return True  # セル編集モードにする。
	def mouseReleased(self, enhancedmouseevent):
		return True  # シングルクリックでFalseを返すとセル選択範囲の決定の状態になってどうしようもなくなる。
	def disposing(self, eventobject):  # ドキュメントを閉じる時でも呼ばれない。
		self.subj.removeEnhancedMouseClickHandler(self)	
def createDialog(xscriptcontext, enhancedmouseevent):	
	
# 	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	doc = xscriptcontext.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
	dialogpoint = getDialogPoint(doc, enhancedmouseevent)  # クリックした位置をma単位でPointで取得。
	if dialogpoint:
		frame = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
		containerwindow = frame.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
		toolkit = containerwindow.getToolkit()  # ピアからツールキットを取得。  
		m = 6  # コントロール間の間隔
		grid = {"PositionX": m, "PositionY": m, "Width": 100, "Height": 50, "ShowRowHeader": False, "ShowColumnHeader": False, "SelectionModel": MULTI, "VScroll": True}  # グリッドコントロールの基本プロパティ。
		textbox = {"PositionX": m, "PositionY": YHeight(grid, m), "Height": 12}  # テクストボックスコントロールの基本プロパティ。
		button = {"PositionY": textbox["PositionY"]-1, "Width": 23, "Height":textbox["Height"]+2, "PushButtonType": 2}  # ボタンの基本プロパティ。PushButtonTypeの値はEnumではエラーになる。VerticalAlignではtextboxと高さが揃わない。
		controldialog =  {"PositionX": dialogpoint.X, "PositionY": dialogpoint.Y, "Width": grid["PositionX"]+grid["Width"]+m, "Title": "Grid Example", "Name": "controldialog", "Moveable": True}  # コントロールダイアログの基本プロパティ。幅は右端のコントロールから取得。高さは最後に設定する。
		dialog, addControl = dialogCreator(ctx, smgr, controldialog)  # コントロールダイアログの作成。
		menulistener = MenuListener()  # コンテクストメニューのリスナー。
		mouselistener = MouseListener(doc, menulistener, menuCreator(ctx, smgr))
		gridselectionlistener = GridSelectionListener()
		gridcontrol1 = addControl("Grid", grid, {"addMouseListener": mouselistener, "addSelectionListener": gridselectionlistener})  # グリッドコントロールの取得。
		gridmodel = gridcontrol1.getModel()  # グリッドコントロールモデルの取得。
		gridcolumn = gridmodel.getPropertyValue("ColumnModel")  # DefaultGridColumnModel
		column0 = gridcolumn.createColumn()  # 列の作成。
		column0.ColumnWidth = 50  # 列幅。
		gridcolumn.addColumn(column0)  # 列を追加。
		column1 = gridcolumn.createColumn()  # 列の作成。
		column1.ColumnWidth = grid["Width"] - column0.ColumnWidth  #  列幅。列の合計がグリッドコントロールの幅に一致するようにする。
		gridcolumn.addColumn(column1)  # 列を追加。	
		griddata = gridmodel.getPropertyValue("GridDataModel")  # GridDataModel
		datarows = getSavedGridRows(doc, "Grid1")  # グリッドコントロールの行をhistoryシートのragenameから取得する。	
		now = datetime.now()  # 現在の日時を取得。
		d = now.date().isoformat()
		t = now.time().isoformat().split(".")[0]	
		if datarows:  # 行のリストが取得出来た時。
			griddata.insertRows(0, ("",)*len(datarows), datarows)  # グリッドに行を挿入。
		else:
			griddata.addRow("", (t, d))  # 現在の行を入れる。
		textbox1, textbox2 = [textbox.copy() for dummy in range(2)]
		textbox1["Width"] = 34
		textbox1["Text"] = t
		textbox2["PositionX"] = XWidth(textbox1) 
		textbox2["Width"] = 42
		textbox2["Text"] = d
		addControl("Edit", textbox1)  
		addControl("Edit", textbox2, {"addMouseListener": mouselistener})  
		button["Label"] = "~Close"
		button["PositionX"] = XWidth(textbox2) 
		addControl("Button", button, {"addMouseListener": mouselistener})  
		dialog.getModel().setPropertyValue("Height", YHeight(button, m))  # コントロールダイアログの高さを設定。
		dialog.createPeer(toolkit, containerwindow)  # ダイアログを描画。親ウィンドウを渡す。ノンモダルダイアログのときはNone(デスクトップ)ではフリーズする。Stepを使うときはRoadmap以外のコントロールが追加された後にピアを作成しないとStepが重なって表示される。
		menulistener.setDialog(dialog)
		dialogframe = showModelessly(ctx, smgr, frame, dialog)  # ノンモダルダイアログを表示。
		frameactionlistener = FrameActionListener()  # FrameActionListener。フレームがアクティブでなくなった時に閉じるため。
		dialogframe.addFrameActionListener(frameactionlistener)  # FrameActionListenerをダイアログフレームに追加。
		args = mouselistener, gridselectionlistener
		dialogframe.addCloseListener(CloseListener(args))  # CloseListener。ノンモダルダイアログのリスナー削除用。	
def XWidth(props, m=0):  # 左隣のコントロールからPositionXを取得。mは間隔。
	return props["PositionX"] + props["Width"] + m  	
def YHeight(props, m=0):  # 上隣のコントロールからPositionYを取得。mは間隔。
	return props["PositionY"] + props["Height"] + m
def getDialogPoint(doc, enhancedmouseevent):  # クリックした位置をPointで返す。但し、一部しか見えてないセルの場合はNoneが返る。
	controller = doc.getCurrentController()  # 現在のコントローラを取得。
	frame = controller.getFrame()  # フレームを取得。
	containerwindow = frame.getContainerWindow()  # コンテナウィドウの取得。
	framepointonscreen = containerwindow.getAccessibleContext().getAccessibleParent().getAccessibleContext().getLocationOnScreen()  # フレームの左上角の点（画面の左上角が原点)。
	componentwindow = frame.getComponentWindow()  # コンポーネントウィンドウを取得。
	border = controller.getBorder()  # 行ヘッダの幅と列ヘッダの高さの取得のため。
	accessiblecontext = componentwindow.getAccessibleContext()  # コンポーネントウィンドウのAccessibleContextを取得。
	for i in range(accessiblecontext.getAccessibleChildCount()):  # 子AccessibleContextについて。
		childaccessiblecontext = accessiblecontext.getAccessibleChild(i).getAccessibleContext()  # 子AccessibleContextのAccessibleContext。
		if childaccessiblecontext.getAccessibleRole()==51:  # SCROLL_PANEの時。
			for j in range(childaccessiblecontext.getAccessibleChildCount()):  # 孫AccessibleContextについて。 
				grandchildaccessiblecontext = childaccessiblecontext.getAccessibleChild(j).getAccessibleContext()  # 孫AccessibleContextのAccessibleContext。
				if grandchildaccessiblecontext.getAccessibleRole()==84:  # DOCUMENT_SPREADSHEETの時。これが枠。
					bounds = grandchildaccessiblecontext.getBounds()  # 枠の位置と大きさを取得(SCROLL_PANEの左上角が原点)。
					if bounds.X==border.Left and bounds.Y==border.Top:  # SCROLL_PANEに対する相対座標が行ヘッダと列ヘッダと一致する時は左上枠。
						for k, subcontroller in enumerate(controller):  # 各枠のコントローラについて。インデックスも取得する。
							cellrange = subcontroller.getReferredCells()  # 見えているセル範囲を取得。一部しかみえていないセルは含まれない。
							if len(cellrange.queryIntersection(enhancedmouseevent.Target.getRangeAddress())):  # ターゲットが含まれるセル範囲コレクションが返る時その枠がクリックした枠。「ウィンドウの分割」では正しいiは必ずしも取得できない。
								sourcepointonscreen =  grandchildaccessiblecontext.getLocationOnScreen()  # 左上枠の左上角の点を取得(画面の左上角が原点)。
								if k==1:  # 左下枠の時。
									sourcepointonscreen = Point(X=sourcepointonscreen.X, Y=sourcepointonscreen.Y+bounds.Height)
								elif k==2:  # 右上枠の時。
									sourcepointonscreen = Point(X=sourcepointonscreen.X+bounds.Width, Y=sourcepointonscreen.Y)
								elif k==3:  # 右下枠の時。
									sourcepointonscreen = Point(X=sourcepointonscreen.X+bounds.Width, Y=sourcepointonscreen.Y+bounds.Height)
								x = sourcepointonscreen.X + enhancedmouseevent.X - framepointonscreen.X  # ウィンドウの左上角からの相対Xの取得。
								y = sourcepointonscreen.Y + enhancedmouseevent.Y - framepointonscreen.Y  # ウィンドウの左上角からの相対Yの取得。
								return componentwindow.convertPointToLogic(Point(X=x, Y=y), MeasureUnit.APPFONT)  # ピクセル単位をma単位に変換。
class FrameActionListener(unohelper.Base, XFrameActionListener):
	def frameAction(self, frameactionevent):
		if frameactionevent.Action==FRAME_UI_DEACTIVATING:  # フレームがアクティブでなくなった時。TopWindowListenerのwindowDeactivated()だとウィンドウタイトルバーをクリックしただけで発火してしまう。
			frameactionevent.Frame.removeFrameActionListener(self)  # フレームにつけたリスナーを除去。
			frameactionevent.Frame.close(True)
	def disposing(self, eventobject):
		eventobject.Source.removeFrameActionListener(self)
class CloseListener(unohelper.Base, XCloseListener):  # ノンモダルダイアログのリスナー削除用。
	def __init__(self, args):
		self.args = args
	def queryClosing(self, eventobject, getsownership):
		mouselistener, gridselectionlistener = self.args
		doc, menulistener, gridpopupmenu, editpopupmenu, buttonpopupmenu = mouselistener.args
		dialog = menulistener.args[0]
		gridcontrol = dialog.getControl("Grid1")	
		saveGridRows(doc, gridcontrol, "Grid1")
		gridpopupmenu.removeMenuListener(menulistener)
		editpopupmenu.removeMenuListener(menulistener)
		buttonpopupmenu.removeMenuListener(menulistener)
		gridcontrol.removeSelectionListener(gridselectionlistener)
		gridcontrol.removeMouseListener(mouselistener)
		dialog.getControl("Edit2").removeMouseListener(mouselistener)
		dialog.getControl("Button1").removeMouseListener(mouselistener)
		eventobject.Source.removeCloseListener(self)
	def notifyClosing(self, eventobject):
		pass
	def disposing(self, eventobject):  
		eventobject.Source.removeCloseListener(self)
def saveGridRows(doc, gridcontrol, rangename):  # グリッドコントロールの行をhistoryシートのragenameに保存する。		
	griddatamodel = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModel
	datarows = [griddatamodel.getRowData(i) for i in range(griddatamodel.RowCount)]  # グリッドコントロールの行のリストを取得。
	namedranges = doc.getPropertyValue("NamedRanges")  # ドキュメントのNamedRangesを取得。
	if not rangename in namedranges:  # 名前がない時。名前は重複しているとエラーになる。
		sheets = doc.getSheets()  # シートコレクションを取得。
		sheetname = "history"  # 履歴シート名。
		if not sheetname in sheets:  # 履歴シートがない時。
			sheets.insertNewByName(sheetname, len(sheets))   # 履歴シートを挿入。同名のシートがあるとRuntimeExceptionがでる。
		sheet = sheets[sheetname]  # 履歴シートを取得。
		sheet.setPropertyValue("IsVisible", False)  # 非表示シートにする。
		emptyranges = sheet[:, :2].queryEmptyCells()  # 2列目までの最初の空セル範囲コレクションを取得。
		if len(emptyranges):  # セル範囲コレクションが取得出来た時。
			emptyrange = emptyranges[0]  # 最初のセル範囲を取得。
			emptyrange[0, 0].setString(rangename)
			namedranges.addNewByName(rangename, emptyrange[0, 1].getPropertyValue("AbsoluteName"), emptyrange[0, 1].getCellAddress(), 0)  # 2列目のセルに名前を付ける。名前、式(相対アドレス)、原点となるセル、NamedRangeFlag
	namedranges[rangename].getReferredCells().setString(json.dumps(datarows,  ensure_ascii=False))  # Grid1という名前のセルに文字列でリストを出力する。
def getSavedGridRows(doc, rangename):  # グリッドコントロールの行をhistoryシートのragenameから取得する。	
	namedranges = doc.getPropertyValue("NamedRanges")  # ドキュメントのNamedRangesを取得。	
	if rangename in namedranges:  # 名前がある時。
		txt = namedranges[rangename].getReferredCells().getString()  # 名前が参照しているセルから文字列を取得。
		if txt:
			try:
				return json.loads(txt)
			except json.JSONDecodeError:
				pass
	return None  # 保存された行が取得できない時はNoneを返す。
class MouseListener(unohelper.Base, XMouseListener):  
	def __init__(self, doc, menulistener, createMenu): 
		items = ("~Cut", 0, {"setCommand": "cut"}),\
			("Cop~y", 0, {"setCommand": "copy"}),\
			("~Paste Above", 0, {"setCommand": "pasteabove"}),\
			("P~aste Below", 0, {"setCommand": "pastebelow"}),\
			(),\
			("~Delete Selected Rows", 0, {"setCommand": "delete"})  # グリッドコントロールにつける右クリックメニュー。
		gridpopupmenu = createMenu("PopupMenu", items, {"addMenuListener": menulistener})  # 右クリックでまず呼び出すポップアップメニュー。  
		items = ("~Now", 0, {"setCommand": "now"}),  # テキストボックスコントロールにつける右クリックメニュー。
		editpopupmenu = createMenu("PopupMenu", items, {"addMenuListener": menulistener})  # 右クリックでまず呼び出すポップアップメニュー。  	
		items = ("~Resore", 0, {"setCommand": "restore"}),\
			(),\
			("~Add", 0, {"setCommand": "add"}),\
			("~Sort", 0, {"setCommand": "sort"})  # ボタンコントロールにつける右クリックメニュー。
		buttonpopupmenu = createMenu("PopupMenu", items, {"addMenuListener": menulistener})  # 右クリックでまず呼び出すポップアップメニュー。  
		self.args = doc, menulistener, gridpopupmenu, editpopupmenu, buttonpopupmenu
	def mousePressed(self, mouseevent):  # グリッドコントロールをクリックした時。
		doc, menulistener, gridpopupmenu, editpopupmenu, buttonpopupmenu = self.args
		name = mouseevent.Source.getModel().getPropertyValue("Name")
		if name=="Grid1":  # グリッドコントロールの時。
			gridcontrol = mouseevent.Source  # グリッドコントロールを取得。
			if mouseevent.Buttons==MouseButton.LEFT and mouseevent.ClickCount==2:  # ダブルクリックの時。
				selection = doc.getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
				if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択オブジェクトがセルの時。
					griddata = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。
					rowdata = griddata.getRowData(gridcontrol.getCurrentRow())  # グリッドコントロールで選択している行のすべての列をタプルで取得。
					cellcursor = selection.getSpreadsheet().createCursorByRange(selection)  # 選択範囲のセルカーサーを取得。
					cellcursor.collapseToSize(len(rowdata), 1)  # (列、行)で指定。セルカーサーの範囲をrowdataに合せる。
					menulistener.undo = cellcursor, cellcursor.getDataArray()  # undoのためにセルカーサーとその値を取得する。
					cellcursor.setDataArray((rowdata,))  # セルカーサーにrowdataを代入。代入できるのは整数(int、ただしboolを除く)か文字列のみ。
			elif mouseevent.PopupTrigger:  # 右クリックの時。
				rowindex = gridcontrol.getRowAtPoint(mouseevent.X, mouseevent.Y)  # クリックした位置の行インデックスを取得。該当行がない時は-1が返ってくる。
				if rowindex>-1:  # クリックした位置に行が存在する時。
					flg = True  # Pasteメニューを表示させるフラグ。
					if not gridcontrol.isRowSelected(rowindex):  # クリックした位置の行が選択状態でない時。
						gridcontrol.deselectAllRows()  # 行の選択状態をすべて解除する。
						gridcontrol.selectRow(rowindex)  # 右クリックしたところの行を選択する。
					rows = gridcontrol.getSelectedRows()  # 選択行インデックスを取得。
					rowcount = len(rows)  # 選択行数を取得。
					if rowcount>1 or not menulistener.rowdata:  # 複数行を選択している時または保存データがない時。
						flg = False  # Pasteメニューを表示しない。
					gridpopupmenu.enableItem(3, flg)  
					gridpopupmenu.enableItem(4, flg)  			
					pos = Rectangle(mouseevent.X, mouseevent.Y, 0, 0)  # ポップアップメニューを表示させる起点。
					gridpopupmenu.execute(gridcontrol.getPeer(), pos, PopupMenuDirection.EXECUTE_DEFAULT)  # ポップアップメニューを表示させる。引数は親ピア、位置、方向	
		elif name=="Edit2":  # テキストボックスコントロールの時。			
			if mouseevent.Buttons==MouseButton.LEFT and mouseevent.ClickCount==2:  # ダブルクリックの時。テキストボックスコントロールでは右クリックはカスタマイズ出来ない。
				editcontrol = mouseevent.Source  # テキストボックスコントロールを取得。
				pos = Rectangle(mouseevent.X, mouseevent.Y, 0, 0)  # ポップアップメニューを表示させる起点。
				editpopupmenu.execute(editcontrol.getPeer(), pos, PopupMenuDirection.EXECUTE_DEFAULT)  # ポップアップメニューを表示させる。引数は親ピア、位置、方向						
		elif name=="Button1":  # ボタンコントロールの時。
			if mouseevent.PopupTrigger:  # 右クリックの時。
				flg = False  # Undoメニューを表示させるフラグ。
				if menulistener.undo:  # Undoデータがある時。
					cellcursor = menulistener.undo[0]  # Undoするセルカーサーを取得。
					activesheetname = doc.getCurrentController().getActiveSheet().getName()
					if activesheetname==cellcursor.getSpreadsheet().getName():  # Undoデータと同じシートの時。
						flg = True
				buttonpopupmenu.enableItem(1, flg)  # Undoメニューを表示する。
				buttoncontrol = mouseevent.Source  # ボタンコントロールを取得。
				pos = Rectangle(mouseevent.X, mouseevent.Y, 0, 0)  # ポップアップメニューを表示させる起点。
				buttonpopupmenu.execute(buttoncontrol.getPeer(), pos, PopupMenuDirection.EXECUTE_DEFAULT)  # ポップアップメニューを表示させる。引数は親ピア、位置、方向					
	def mouseReleased(self, mouseevent):
		pass
	def mouseEntered(self, mouseevent):
		pass
	def mouseExited(self, mouseevent):
		pass
	def disposing(self, eventobject):
		eventobject.Source.removeMouseListener(self)
class MenuListener(unohelper.Base, XMenuListener):
	def __init__(self):  # グリッドコントロールはこの時点でまだdialogに追加されていない。ピアも作成されていない。
		self.rowdata = None
		self.undo = None  # undo用データ。
	def setDialog(self, dialog):  # グリッドコントロールとピアが作成されてから実行する。
		peer = dialog.getPeer()  # ピアを取得。
		toolkit = peer.getToolkit()  # ピアからツールキットを取得。 	
		gridcontrol = dialog.getControl("Grid1")  # グリッドコントロールを取得。	
		griddata = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。		
		self.args = dialog, peer, toolkit, gridcontrol, griddata  # dialogはCloseListener内で使うので最初に置かないといけない。
	def itemHighlighted(self, menuevent):
		pass
	def itemSelected(self, menuevent):  # PopupMenuの項目がクリックされた時。どこのコントロールのメニューかを知る方法はない。
		dialog, peer, toolkit, gridcontrol, griddata = self.args
		cmd = menuevent.Source.getCommand(menuevent.MenuId)
		selectedrows = gridcontrol.getSelectedRows()  # 選択行インデックスのタプルを取得。
		if cmd in ("cut", "copy", "pasteabove", "pastebelow", "delete"):  # グリッドコントロールのコンテクストメニュー。
			if cmd=="cut":  # 選択行のデータを取得してその行を削除する。
				self.rowdata = [griddata.getRowData(r) for r in selectedrows]  # 選択行のデータを取得。
				[griddata.removeRow(r) for r in selectedrows]  # 選択行を削除。
			elif cmd=="copy":  # 選択行のデータを取得する。  
				self.rowdata = [griddata.getRowData(r) for r in selectedrows]  # 選択行のデータを取得。
			elif cmd=="pasteabove":  # 行を選択行の上に挿入。 
				insertRows(gridcontrol, griddata, selectedrows, 0, self.rowdata)
			elif cmd=="pastebelow":  # 空行を選択行の下に挿入。  
				insertRows(gridcontrol, griddata, selectedrows, 1, self.rowdata)
			elif cmd=="delete":  # 選択行を削除する。  
				msg = "Delete selected row(s)?"
				msgbox = toolkit.createMessageBox(peer, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO, "Delete", msg)
				if msgbox.execute()==MessageBoxResults.YES:					
					[griddata.removeRow(r) for r in selectedrows]  # 選択行を削除。
		elif cmd in ("add", "restore", "sort"):  # ボタンコントロールのコンテクストメニュー。
			if cmd=="add":
				t = dialog.getControl("Edit1").getText()
				d = dialog.getControl("Edit2").getText()			
				if not selectedrows:  # 選択行がない時。
					selectedrows = griddata.RowCount-1,  # 最終行インデックスを選択していることにする。
				insertRows(gridcontrol, griddata, selectedrows, 1, ((t, d),))  # 選択行の下に行を挿入する。
			elif cmd=="sort":
				msg = "Sort in ascending order?"
				msgbox = toolkit.createMessageBox(peer, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO, "Sort", msg)
				if msgbox.execute()==MessageBoxResults.YES:				
					griddata.sortByColumn(0, True)
			elif cmd=="restore":
				cellcursor, datarows = self.undo  # datarowsは1行しかないはず。
				stringaddress = cellcursor.getPropertyValue("AbsoluteName").split(".")[1].replace("$", "")  # 前回入力した範囲の文字列アドレスを取得。
				current = " ".join(cellcursor.getDataArray()[0])
				restored = " ".join(datarows[0])
				msg = """Restore the Value of {}?
Current: {}
  After: {}""".format(stringaddress, current, restored)
				msgbox = toolkit.createMessageBox(peer, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO, "Undo", msg)
				if msgbox.execute()==MessageBoxResults.YES:
					cellcursor.setDataArray(datarows)
		elif cmd in ("now",):
			now = datetime.now()  # 現在の日時を取得。
			dialog.getControl("Edit2").setText(now.date().isoformat())  # テキストボックスコントロールに入力。
			dialog.getControl("Edit1").setText(now.time().isoformat().split(".")[0])  # テキストボックスコントロールに入力。			
	def itemActivated(self, menuevent):
		pass
	def itemDeactivated(self, menuevent):
		pass   
	def disposing(self, eventobject):
		eventobject.Source.removeMenuListener(self)
class GridSelectionListener(unohelper.Base, XGridSelectionListener):
	def selectionChanged(self, gridselectionevent):  # 行を追加した時も発火する。
		gridcontrol = gridselectionevent.Source
		selectedrows = gridselectionevent.SelectedRowIndexes  # 行がないグリッドコントロールに行が追加されたときは負の値が入ってくる。
		if selectedrows:  # 選択行がある時。
			rowdata = gridcontrol.getModel().getPropertyValue("GridDataModel").getRowData(gridselectionevent.SelectedRowIndexes[0])  # 選択行の最初の行のデータを取得。
			dialog = gridcontrol.getContext()
			dialog.getControl("Edit1").setText(rowdata[0])
			dialog.getControl("Edit2").setText(rowdata[1])
	def disposing(self, eventobject):
		eventobject.Source.removeSelectionListener(self)	
def insertRows(gridcontrol, griddata, selectedrows, position, datarows):  # positionは0の時は選択行の上に挿入、1で下に挿入。
	c = len(datarows)  # 行数を取得。
	griddata.insertRows(selectedrows[0]+position, ("", )*c, datarows)  # 行を挿入。
	gridcontrol.deselectAllRows()  # 行の選択状態をすべて解除する。
	gridcontrol.selectRow(selectedrows[0]+position)  # 挿入した行の最初の行を選択する。	
def menuCreator(ctx, smgr):  #  メニューバーまたはポップアップメニューを作成する関数を返す。
	def createMenu(menutype, items, attr=None):  # menutypeはMenuBarまたはPopupMenu、itemsは各メニュー項目の項目名、スタイル、適用するメソッドのタプルのタプル、attrは各項目に適用する以外のメソッド。
		if attr is None:
			attr = {}
		menu = smgr.createInstanceWithContext("com.sun.star.awt.{}".format(menutype), ctx)
		for i, item in enumerate(items, start=1):  # 各メニュー項目について。
			if item:
				if len(item) > 2:  # タプルの要素が3以上のときは3番目の要素は適用するメソッドの辞書と考える。
					item = list(item)
					attr[i] = item.pop()  # メニュー項目のIDをキーとしてメソッド辞書に付け替える。
				menu.insertItem(i, *item, i-1)  # ItemId, Text, ItemSytle, ItemPos。ItemIdは1から始まり区切り線(欠番)は含まない。ItemPosは0から始まり区切り線を含む。
			else:  # 空のタプルの時は区切り線と考える。
				menu.insertSeparator(i-1)  # ItemPos
		if attr:  # メソッドの適用。
			for key, val in attr.items():  # keyはメソッド名あるいはメニュー項目のID。
				if isinstance(val, dict):  # valが辞書の時はkeyは項目ID。valはcreateMenu()の引数のitemsであり、itemsの３番目の要素にキーをメソッド名とする辞書が入っている。
					for method, arg in val.items():  # 辞書valのキーはメソッド名、値はメソッドの引数。
						if method in ("checkItem", "enableItem", "setCommand", "setHelpCommand", "setHelpText", "setTipHelpText"):  # 第1引数にIDを必要するメソッド。
							getattr(menu, method)(key, arg)
						else:
							getattr(menu, method)(arg)
				else:
					getattr(menu, key)(val)
		return menu
	return createMenu
def showModelessly(ctx, smgr, parentframe, dialog):  # ノンモダルダイアログにする。オートメーションでは動かない。ノンモダルダイアログではフレームに追加しないと閉じるボタンが使えない。
	frame = smgr.createInstanceWithContext("com.sun.star.frame.Frame", ctx)  # 新しいフレームを生成。
	frame.initialize(dialog.getPeer())  # フレームにコンテナウィンドウを入れる。	
	frame.setName(dialog.getModel().getPropertyValue("Name"))  # フレーム名をダイアログモデル名から取得（一致させる必要性はない）して設定。ｽﾍﾟｰｽは不可。
	parentframe.getFrames().append(frame)  # 新しく作ったフレームを既存のフレームの階層に追加する。 
	dialog.setVisible(True)  # ダイアログを見えるようにする。   
	return frame  # フレームにリスナーをつけるときのためにフレームを返す。
def dialogCreator(ctx, smgr, dialogprops):  # ダイアログと、それにコントロールを追加する関数を返す。まずダイアログモデルのプロパティを取得。
	dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)  # ダイアログの生成。
	if "PosSize" in dialogprops:  # コントロールモデルのプロパティの辞書にPosSizeキーがあるときはピクセル単位でコントロールに設定をする。
		dialog.setPosSize(dialogprops.pop("PositionX"), dialogprops.pop("PositionY"), dialogprops.pop("Width"), dialogprops.pop("Height"), dialogprops.pop("PosSize"))  # ダイアログモデルのプロパティで設定すると単位がMapAppになってしまうのでコントロールに設定。
	dialogmodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)  # ダイアログモデルの生成。
	dialogmodel.setPropertyValues(tuple(dialogprops.keys()), tuple(dialogprops.values()))  # ダイアログモデルのプロパティを設定。
	dialog.setModel(dialogmodel)  # ダイアログにダイアログモデルを設定。
	dialog.setVisible(False)  # 描画中のものを表示しない。
	def addControl(controltype, props, attrs=None):  # props: コントロールモデルのプロパティ、attr: コントロールの属性。
		control = None
		items, currentitemid = None, None
		if controltype == "Roadmap":  # Roadmapコントロールのとき、Itemsはダイアログモデルに追加してから設定する。そのときはCurrentItemIDもあとで設定する。
			if "Items" in props:  # Itemsはダイアログモデルに追加されてから設定する。
				items = props.pop("Items")
				if "CurrentItemID" in props:  # CurrentItemIDはItemsを追加されてから設定する。
					currentitemid = props.pop("CurrentItemID")
		if "PosSize" in props:  # コントロールモデルのプロパティの辞書にPosSizeキーがあるときはピクセル単位でコントロールに設定をする。
			if controltype=="Grid":
				control = smgr.createInstanceWithContext("com.sun.star.awt.grid.UnoControl{}".format(controltype), ctx)  # コントロールを生成。
			else:	
				control = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl{}".format(controltype), ctx)  # コントロールを生成。
			control.setPosSize(props.pop("PositionX"), props.pop("PositionY"), props.pop("Width"), props.pop("Height"), props.pop("PosSize"))  # ピクセルで指定するために位置座標と大きさだけコントロールで設定。
			controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
			control.setModel(controlmodel)  # コントロールにコントロールモデルを設定。
			dialog.addControl(props["Name"], control)  # コントロールをコントロールコンテナに追加。
		else:  # Map AppFont (ma)のときはダイアログモデルにモデルを追加しないと正しくピクセルに変換されない。
			controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
			dialogmodel.insertByName(props["Name"], controlmodel)  # ダイアログモデルにモデルを追加するだけでコントロールも作成される。
		if items is not None:  # コントロールに追加されたRoadmapモデルにしかRoadmapアイテムは追加できない。
			for i, j in enumerate(items):  # 各Roadmapアイテムについて
				item = controlmodel.createInstance()
				item.setPropertyValues(("Label", "Enabled"), j)
				controlmodel.insertByIndex(i, item)  # IDは0から整数が自動追加される
			if currentitemid is not None:  #Roadmapアイテムを追加するとそれがCurrentItemIDになるので、Roadmapアイテムを追加してからCurrentIDを設定する。
				controlmodel.setPropertyValue("CurrentItemID", currentitemid)
		if control is None:  # コントロールがまだインスタンス化されていないとき
			control = dialog.getControl(props["Name"])  # コントロールコンテナに追加された後のコントロールを取得。
		if attrs is not None:  # Dialogに追加したあとでないと各コントロールへの属性は追加できない。
			for key, val in attrs.items():  # メソッドの引数がないときはvalをNoneにしている。
				if val is None:
					getattr(control, key)()
				else:
					getattr(control, key)(val)
		return control  # 追加したコントロールを返す。
	def _createControlModel(controltype, props):  # コントロールモデルの生成。
		if not "Name" in props:
			props["Name"] = _generateSequentialName(controltype)  # Nameがpropsになければ通し番号名を生成。
		if controltype=="Grid":
			controlmodel = dialogmodel.createInstance("com.sun.star.awt.grid.UnoControl{}Model".format(controltype))  # コントロールモデルを生成。UnoControlDialogElementサービスのためにUnoControlDialogModelからの作成が必要。
		else:	
			controlmodel = dialogmodel.createInstance("com.sun.star.awt.UnoControl{}Model".format(controltype))  # コントロールモデルを生成。UnoControlDialogElementサービスのためにUnoControlDialogModelからの作成が必要。
		if props:
			values = props.values()  # プロパティの値がタプルの時にsetProperties()でエラーが出るのでその対応が必要。
			if any(map(isinstance, values, [tuple]*len(values))):
				[setattr(controlmodel, key, val) for key, val in props.items()]  # valはリストでもタプルでも対応可能。XMultiPropertySetのsetPropertyValues()では[]anyと判断されてタプルも使えない。
			else:
				controlmodel.setPropertyValues(tuple(props.keys()), tuple(values))
		return controlmodel
	def _generateSequentialName(controltype):  # コントロールの連番名の作成。
		i = 1
		flg = True
		while flg:
			name = "{}{}".format(controltype, i)
			flg = dialog.getControl(name)  # 同名のコントロールの有無を判断。
			i += 1
		return name
	return dialog, addControl  # コントロールコンテナとそのコントロールコンテナにコントロールを追加する関数を返す。
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
