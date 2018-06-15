#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from datetime import datetime
import json
from com.sun.star.accessibility import AccessibleRole  # 定数
from com.sun.star.awt import XActionListener
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.awt import XMenuListener, XItemListener
from com.sun.star.awt import XMouseListener
from com.sun.star.awt import MouseButton, PosSize  # 定数
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
from com.sun.star.beans import NamedValue  # Struct
from com.sun.star.awt import XWindowListener
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
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	doc = xscriptcontext.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
	dialogpoint = getDialogPoint(doc, enhancedmouseevent)  # クリックした位置をma単位でPointで取得。一部しか表示されていないセルのときはNoneが返る。
	if dialogpoint:
		docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
		containerwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
		toolkit = containerwindow.getToolkit()  # ピアからツールキットを取得。  
		m = 6  # コントロール間の間隔
		h = 12
		gridprops = {"PositionX": m, "PositionY": m, "Width": 104, "Height": 50, "ShowRowHeader": False, "ShowColumnHeader": False, "SelectionModel": MULTI, "VScroll": True}  # グリッドコントロールの基本プロパティ。
		radiobuttonprops = {"PositionX": 0, "PositionY": YHeight(gridprops), "Height": h, "Width": 25, "State": 0}  # ラジオボタンコントロールの基本プロパティ。
		textboxprops = {"PositionX": m, "PositionY": YHeight(radiobuttonprops, 2), "Width": 0, "Height": h}  # テクストボックスコントロールの基本プロパティ。
		buttonprops = {"PositionX": 0, "PositionY": YHeight(textboxprops, 4), "Width": 30, "Height":h+2}  # ボタンの基本プロパティ。PushButtonTypeの値はEnumではエラーになる。VerticalAlignではtextboxと高さが揃わない。
		controlcontainerprops = {"PositionX": 0, "PositionY": 0, "Width": gridprops["PositionX"]+gridprops["Width"]+m, "Height": 500, "Moveable": True, "BackgroundColor": 0xF0F0F0}  # コントロールコンテナの基本プロパティ。幅は右端のコントロールから取得。高さはコントロール追加後に最後に設定し直す。
		maTopx = createConverters(containerwindow)  # ma単位をピクセルに変換する関数を取得。
		controlcontainer, addControl = controlcontainerMaCreator(ctx, smgr, maTopx, controlcontainerprops)  # コントロールコンテナの作成。		
		menulistener = MenuListener(controlcontainer)  # コンテクストメニューのリスナー。
		mouselistener = MouseListener(doc, menulistener, menuCreator(ctx, smgr))
		gridselectionlistener = GridSelectionListener()
		gridcontrol1 = addControl("Grid", gridprops.copy(), {"addMouseListener": mouselistener, "addSelectionListener": gridselectionlistener})  # グリッドコントロールの取得。gridは他のコントロールの設定に使うのでコピーを渡す。
		gridmodel = gridcontrol1.getModel()  # グリッドコントロールモデルの取得。
		gridcolumn = gridmodel.getPropertyValue("ColumnModel")  # DefaultGridColumnModel
		column0 = gridcolumn.createColumn()  # 列の作成。
		gridcolumn.addColumn(column0)  # 列を追加。
		griddata = gridmodel.getPropertyValue("GridDataModel")  # GridDataModel
		datarows = getSavedGridRows(doc, "Grid1")  # グリッドコントロールの行をhistoryシートのragenameから取得する。	

		now = datetime.now()  # 現在の日時を取得。
		t = now.isoformat()

		if datarows:  # 行のリストが取得出来た時。
			griddata.insertRows(0, ("",)*len(datarows), datarows)  # グリッドに行を挿入。
		else:
			griddata.addRow("", (t,))  # 現在の行を入れる。
		radiobuttons = [radiobuttonprops.copy() for dummy in range(4)]
		radiobutton1, radiobutton2, radiobutton3, radiobutton4 = radiobuttons
		radiobutton1["Label"] = "~昇順"
		radiobutton2["Label"] = "~降順"
		radiobutton2["PositionX"] = XWidth(radiobutton1) 
		radiobutton3["Label"] = "~頻度順"
		radiobutton3["PositionX"] = XWidth(radiobutton2) 
		radiobutton3["Width"] += 6
		radiobutton4["Label"] = "~使用順"
		radiobutton4["PositionX"] = XWidth(radiobutton3) 
		radiobutton4["Width"] += 6		
		radiobutton4["State"] = 1  # デフォルトの選択。
		itemlistener = ItemListener(controlcontainer)
		[addControl("RadioButton", i, {"addItemListener": itemlistener}) for i in radiobuttons]
		textbox1 = textboxprops.copy()
		textbox1["Width"] = gridprops["Width"]
		textbox1["Text"] = doc.getCurrentSelection().getString()  # セルの文字列を取得してテキストボックスに表示する。
		addControl("Edit", textbox1)  
		button1, button2, button3, button4, button5 = [buttonprops.copy() for dummy in range(2)]
		
		
		
		button1["Label"] = "~セルへ"
		button1["PositionX"] = m + gridprops["Width"] - button1["Width"]
		button2["Label"] = "~クリア"
		button2["PositionX"] = button1["PositionX"] - 2 - button2["Width"]
		button3["Label"] = "~追加"
		button3["PositionX"] = button2["PositionX"] - 2 - button3["Width"]
		button4["Label"] = "~↓"
		button4["Width"] -= 20
		button4["PositionX"] = button3["PositionX"] - 2 - button4["Width"]
		button5["Label"] = "~↑"
		button5["Width"] -= 20
		button5["PositionX"] = button4["PositionX"] - 2 - button5["Width"]		
		
		
		
		actionlistener = ActionListener(xscriptcontext)
		addControl("Button", button1, {"addActionListener": actionlistener, "setActionCommand": "clear"})  
		addControl("Button", button2, {"addActionListener": actionlistener, "setActionCommand": "tocell"})   
		controlcontainer.setPosSize(0, 0, 0, maTopx(0, YHeight(buttonprops, m))[1], PosSize.HEIGHT)  # コントロールダイアログの高さを設定。px単位で設定。
		rectangle = controlcontainer.getPosSize()  # コントロールコンテナのRectangle Structを取得。
		rectangle.X, rectangle.Y = dialogpoint  # 位置を代入。
		taskcreator = smgr.createInstanceWithContext('com.sun.star.frame.TaskCreator', ctx)
		args = NamedValue("PosSize", rectangle), NamedValue("FrameName", "controldialog")  # , NamedValue("MakeVisible", True)  # TaskCreatorで作成するフレームのコンテナウィンドウのプロパティ。
		dialogframe = taskcreator.createInstanceWithArguments(args)  # コンテナウィンドウ付きの新しいフレームの取得。
		dialogwindow = dialogframe.getContainerWindow()  # ダイアログのコンテナウィンドウを取得。
		dialogframe.setTitle("履歴")  # フレームのタイトルを設定。
		docframe.getFrames().append(dialogframe) # 新しく作ったフレームを既存のフレームの階層に追加する。			
		controlcontainer.createPeer(toolkit, dialogwindow) # ウィンドウにコントロールを描画。 
		controlcontainer.setVisible(True)  # コントロールの表示。
		dialogwindow.setVisible(True) # ウィンドウの表示		
		menulistener.setDialog(controlcontainer)
		frameactionlistener = FrameActionListener()  # FrameActionListener。フレームがアクティブでなくなった時に閉じるため。
		dialogframe.addFrameActionListener(frameactionlistener)  # FrameActionListenerをダイアログフレームに追加。
		args = mouselistener, gridselectionlistener, itemlistener
		dialogframe.addCloseListener(CloseListener(args))  # CloseListener。ノンモダルダイアログのリスナー削除用。	
# 		dialogwindow.addWindowListener(WindowListener(controls, minsizes)) # setVisible(True)でも呼び出されるので、その後でリスナーを追加する。
def XWidth(props, m=0):  # 左隣のコントロールからPositionXを取得。mは間隔。
	return props["PositionX"] + props["Width"] + m  	
def YHeight(props, m=0):  # 上隣のコントロールからPositionYを取得。mは間隔。
	return props["PositionY"] + props["Height"] + m
def getDialogPoint(doc, enhancedmouseevent):  # クリックした位置x yのタプルで返す。但し、一部しか見えてないセルの場合はNoneが返る。
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
								return x, y
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
		mouselistener, gridselectionlistener, itemlistener = self.args
		doc, menulistener, gridpopupmenu, editpopupmenu, buttonpopupmenu = mouselistener.args
		dialog = menulistener.args[0]
		gridcontrol = dialog.getControl("Grid1")	
		saveGridRows(doc, gridcontrol, "Grid1")
		gridpopupmenu.removeMenuListener(menulistener)
		editpopupmenu.removeMenuListener(menulistener)
		buttonpopupmenu.removeMenuListener(menulistener)
		gridcontrol.removeSelectionListener(gridselectionlistener)
		gridcontrol.removeMouseListener(mouselistener)
		[dialog.getControl("RadioButton{}".format(i)).removeItemListener(itemlistener) for i in range(1, 5)]


# 		dialog.getControl("Edit2").removeMouseListener(mouselistener)
# 		dialog.getControl("Button1").removeMouseListener(mouselistener)
		eventobject.Source.removeCloseListener(self)
	def notifyClosing(self, eventobject):
		pass
	def disposing(self, eventobject):  
		eventobject.Source.removeCloseListener(self)

class ItemListener(unohelper.Base, XItemListener): 
	def __init__(self, controlcontainer):
		self.dialog = controlcontainer
	def itemStateChanged(self, itemevent):
		
		
# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
		
		
		controlmodel = itemevent.Source.getModel()  # 発火させたコントロールモデルを取得。
		controllabel = controlmodel.getPropertyValue("Label") # Labelプロパティを取得。		
		
		
		pass
	
	
	
# 		control, dummy_controlmodel, name = eventSource(itemevent)
# 		if name == "CheckBox1":
# 			buttonprops = self.dialog.getControl("Button1")
# 			buttonmodel = buttonprops.getModel()
# 			state = control.getState()
# 			btnenable = True
# 			if state==0 or state==2:
# 				btnenable = False
# 			buttonmodel.setPropertyValue("Enabled", btnenable)
# 		elif name == "ComboBox1":  # コンボボックスは選択した文字列が取得できない。
# 			control.setText(itemevent.Selected)		
	def disposing(self, eventobject):
		eventobject.Source.removeItemListener(self)

class MouseListener(unohelper.Base, XMouseListener):  
	def __init__(self, doc, menulistener, createMenu): 
		items = ("~削除", 0, {"setCommand": "delete"}),  # グリッドコントロールにつける右クリックメニュー。
		gridpopupmenu = createMenu("PopupMenu", items, {"addMenuListener": menulistener})  # 右クリックでまず呼び出すポップアップメニュー。   	
		self.args = doc, menulistener, gridpopupmenu
	def mousePressed(self, mouseevent):  # グリッドコントロールをクリックした時。コントロールモデルにはNameプロパティはない。
		doc, menulistener, gridpopupmenu = self.args
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
				if not gridcontrol.isRowSelected(rowindex):  # クリックした位置の行が選択状態でない時。
					gridcontrol.deselectAllRows()  # 行の選択状態をすべて解除する。
					gridcontrol.selectRow(rowindex)  # 右クリックしたところの行を選択する。		
				pos = Rectangle(mouseevent.X, mouseevent.Y, 0, 0)  # ポップアップメニューを表示させる起点。
				gridpopupmenu.execute(gridcontrol.getPeer(), pos, PopupMenuDirection.EXECUTE_DEFAULT)  # ポップアップメニューを表示させる。引数は親ピア、位置、方向						
	def mouseReleased(self, mouseevent):
		pass
	def mouseEntered(self, mouseevent):
		pass
	def mouseExited(self, mouseevent):
		pass
	def disposing(self, eventobject):
		eventobject.Source.removeMouseListener(self)
class MenuListener(unohelper.Base, XMenuListener):
	def __init__(self, controlcontainer):  # グリッドコントロールはこの時点でまだdialogに追加されていない。ピアも作成されていない。
		self.controlcontainer = controlcontainer
		
# 		self.rowdata = None
# 		self.undo = None  # undo用データ。
# 	def setDialog(self, controlcontainer):  # ピアが作成されてから実行する。
# 		self.controlcontainer = controlcontainer
		
# 		peer = controlcontainer.getPeer()  # ピアを取得。
# 		toolkit = peer.getToolkit()  # ピアからツールキットを取得。 	
# 		gridcontrol = controlcontainer.getControl("Grid1")  # グリッドコントロールを取得。	
		
# 		self.args = peer, toolkit, gridcontrol  # dialogはCloseListener内で使うので最初に置かないといけない。
	def itemHighlighted(self, menuevent):
		pass
	def itemSelected(self, menuevent):  # PopupMenuの項目がクリックされた時。どこのコントロールのメニューかを知る方法はない。
		cmd = menuevent.Source.getCommand(menuevent.MenuId)
		if cmd=="delete":  # 選択行を削除する。  
		
			controlcontainer = self.controlcontainer
		
# 			peer, toolkit, gridcontrol = self.args
			
			
# 		if cmd in ("delete", ):  # グリッドコントロールのコンテクストメニュー。
# 			if cmd=="cut":  # 選択行のデータを取得してその行を削除する。
# 				self.rowdata = [griddata.getRowData(r) for r in selectedrows]  # 選択行のデータを取得。
# 				[griddata.removeRow(r) for r in selectedrows]  # 選択行を削除。
# 			elif cmd=="copy":  # 選択行のデータを取得する。  
# 				self.rowdata = [griddata.getRowData(r) for r in selectedrows]  # 選択行のデータを取得。
# 			elif cmd=="pasteabove":  # 行を選択行の上に挿入。 
# 				insertRows(gridcontrol, griddata, selectedrows, 0, self.rowdata)
# 			elif cmd=="pastebelow":  # 空行を選択行の下に挿入。  
# 				insertRows(gridcontrol, griddata, selectedrows, 1, self.rowdata)
			peer = controlcontainer.getPeer()  # ピアを取得。	
			msg = "選択行を削除しますか?"
			msgbox = peer.getToolkit().createMessageBox(peer, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO, "myRs", msg)
			if msgbox.execute()==MessageBoxResults.YES:		
				gridcontrol = controlcontainer.getControl("Grid1")  # グリッドコントロールを取得。			
				griddata = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。			
				selectedrows = gridcontrol.getSelectedRows()  # 選択行インデックスのタプルを取得。	
				[griddata.removeRow(r) for r in reversed(selectedrows)]  # 選択行を下から削除。		
	def itemActivated(self, menuevent):
		pass
	def itemDeactivated(self, menuevent):
		pass   
	def disposing(self, eventobject):
		eventobject.Source.removeMenuListener(self)

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
# def showModelessly(ctx, smgr, parentframe, dialog):  # ノンモダルダイアログにする。オートメーションでは動かない。ノンモダルダイアログではフレームに追加しないと閉じるボタンが使えない。
# 	frame = smgr.createInstanceWithContext("com.sun.star.frame.Frame", ctx)  # 新しいフレームを生成。
# 	frame.initialize(dialog.getPeer())  # フレームにコンテナウィンドウを入れる。	
# 	frame.setName(dialog.getModel().getPropertyValue("Name"))  # フレーム名をダイアログモデル名から取得（一致させる必要性はない）して設定。ｽﾍﾟｰｽは不可。
# 	parentframe.getFrames().append(frame)  # 新しく作ったフレームを既存のフレームの階層に追加する。 
# 	dialog.setVisible(True)  # ダイアログを見えるようにする。   
# 	return frame  # フレームにリスナーをつけるときのためにフレームを返す。
class WindowListener(unohelper.Base, XWindowListener):
	def __init__(self, controls, minsizes):
		rectangle = controls[0].getPosSize()  # コントロールコンテナの位置と大きさを取得。なぜかwindow.getPosSize()では取得できない。
		self.oldwidth = rectangle.Width  # 変更前の幅を取得しておく。
		self.oldheight = rectangle.Height  # 変更前の高さを取得しておく。
		self.controls = controls
		self.minsizes = minsizes
# 	@enableRemoteDebugging		
	def windowResized(self, windowevent):  # 変化分で計算する。コントロールが表示されないほど小さくされると次から表示がおかしくなる。
		minwidth, minheight = self.minsizes  # サイズ下限を取得。
		newwidth = windowevent.Width if windowevent.Width>minwidth else minwidth  # 変更後のコントロールコンテナの幅を取得。サイズ下限より小さい時は下限値とする。
		newheight = windowevent.Height if windowevent.Height>minheight else minheight  # 変更後のコントロールコンテナの高さを取得。サイズ下限より小さい時は下限値とする。
		self.diff_width = newwidth - self.oldwidth  # 幅変化分
		self.diff_height = newheight -self.oldheight  # 高さ変化分		
		controlcontainer, imagecontrol1, edit1, button1, button2, radiobutton1, radiobutton2, radiobutton3, fixedtext1, fixedtext2 = self.controls  # 再計算するコントロールを取得。
		controlcontainer.setPosSize(0, 0, newwidth, newheight, PosSize.SIZE)  # Flagsで変更する値のみ指定。変更しない値は0(でもなんでもよいはず)。
		self._applyDiff(fixedtext1, PosSize.Y)
		self._applyDiff(fixedtext2, PosSize.Y)
		self._applyDiff(imagecontrol1, PosSize.SIZE)
		self._applyDiff(edit1, PosSize.Y+PosSize.WIDTH)
		self._applyDiff(radiobutton1, PosSize.Y)
		self._applyDiff(radiobutton2, PosSize.Y)
		self._applyDiff(radiobutton3, PosSize.Y)
		self._applyDiff(button1, PosSize.POS)
		self._applyDiff(button2, PosSize.POS)
		imagecontrolrectangle = imagecontrol1.getPosSize()
		fixedtext2.setText("{} x {} px Display Size".format(imagecontrolrectangle.Width, imagecontrolrectangle.Height))
		self.oldwidth = newwidth  # 次の変更前の幅として取得。
		self.oldheight = newheight  # 次の変更前の高さとして取得。		
	def _applyDiff(self, control, possize):  # 第2引数でウィンドウサイズの変化分のみ適用するPosSizeを指定。
		rectangle = control.getPosSize()  # 変更前のコントロールの位置大きさを取得。
		control.setPosSize(rectangle.X+self.diff_width, rectangle.Y+self.diff_height, rectangle.Width+self.diff_width, rectangle.Height+self.diff_height, possize)		
	def windowMoved(self, windowevent):
		pass
	def windowShown(self, eventobject):
		pass
	def windowHidden(self, eventobject):
		pass
	def disposing(self, eventobject):
		pass	

class ActionListener(unohelper.Base, XActionListener):
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
	def actionPerformed(self, actionevent):
		xscriptcontext = self.xscriptcontext
		cmd = actionevent.ActionCommand
		edit1 = actionevent.Source.getContext().getControl("Edit1")
		if cmd=="clear":
			edit1.setText("")
		elif cmd=="tocell":
			doc = xscriptcontext.getDocument()  
			selection = doc.getCurrentSelection()
			if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択オブジェクトがセルの時。
				selection.setString(edit1.getText())
	def disposing(self, eventobject):
		eventobject.Source.removeActionListener(self)
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
			emptyrange[0, 0].setString(rangename)  # 1列目に名前を表示する。
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
class GridSelectionListener(unohelper.Base, XGridSelectionListener):
	def selectionChanged(self, gridselectionevent):  # 行を追加した時も発火する。
		gridcontrol = gridselectionevent.Source
		selectedrows = gridselectionevent.SelectedRowIndexes  # 行がないグリッドコントロールに行が追加されたときは負の値が入ってくる。
		if selectedrows:  # 選択行がある時。
			rowdata = gridcontrol.getModel().getPropertyValue("GridDataModel").getRowData(gridselectionevent.SelectedRowIndexes[0])  # 選択行の最初の行のデータを取得。
			gridcontrol.getContext().getControl("Edit1").setText(rowdata[0])  # テキストボックスに選択行の初行の文字列を代入。
			if len(selectedrows)==1:  # 1行しかない時はまた発火できるように選択を外す。
				gridcontrol.deselectRow(0)
	def disposing(self, eventobject):
		eventobject.Source.removeSelectionListener(self)	
def controlcontainerMaCreator(ctx, smgr, maTopx, containerprops):  # ma単位でコントロールコンテナと、それにコントロールを追加する関数を返す。まずコントロールコンテナモデルのプロパティを取得。UnoControlDialogElementサービスのプロパティは使えない。propsのキーにPosSize、値にPOSSIZEが必要。
	container = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainer", ctx)  # コントロールコンテナの生成。
	container.setPosSize(*maTopx(containerprops.pop("PositionX"), containerprops.pop("PositionY")), *maTopx(containerprops.pop("Width"), containerprops.pop("Height")), PosSize.POSSIZE)
	containermodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainerModel", ctx)  # コンテナモデルの生成。
	containermodel.setPropertyValues(tuple(containerprops.keys()), tuple(containerprops.values()))  # コンテナモデルのプロパティを設定。存在しないプロパティに設定してもエラーはでない。
	container.setModel(containermodel)  # コンテナにコンテナモデルを設定。
	container.setVisible(False)  # 描画中のものを表示しない。
	def addControl(controltype, props, attrs=None):  # props: コントロールモデルのプロパティ、キーにPosSize、値にPOSSIZEが必要。attr: コントロールの属性。
		controlidl = "com.sun.star.awt.grid.UnoControl{}".format(controltype) if controltype=="Grid" else "com.sun.star.awt.UnoControl{}".format(controltype)  # グリッドコントロールだけモジュールが異なる。
		control = smgr.createInstanceWithContext(controlidl, ctx)  # コントロールを生成。
		control.setPosSize(*maTopx(props.pop("PositionX"), props.pop("PositionY")), *maTopx(props.pop("Width"), props.pop("Height")), PosSize.POSSIZE)  # ピクセルで指定するために位置座標と大きさだけコントロールで設定。
		controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
		control.setModel(controlmodel)  # コントロールにコントロールモデルを設定。
		container.addControl(props["Name"], control)  # コントロールをコントロールコンテナに追加。
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
		controlmodelidl = "com.sun.star.awt.grid.UnoControl{}Model".format(controltype) if controltype=="Grid" else "com.sun.star.awt.UnoControl{}Model".format(controltype)
		controlmodel = smgr.createInstanceWithContext(controlmodelidl, ctx) # コントロールモデルを生成。UnoControlDialogElementサービスはない。
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
			flg = container.getControl(name)  # 同名のコントロールの有無を判断。
			i += 1
		return name
	return container, addControl  # コントロールコンテナとそのコントロールコンテナにコントロールを追加する関数を返す。
def createConverters(window):  # ma単位をピクセルに変換する関数を返す。
	def maTopx(x, y):  # maをpxに変換する。
		point = window.convertPointToPixel(Point(X=x, Y=y), MeasureUnit.APPFONT)
		return point.X, point.Y
	return maTopx
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
