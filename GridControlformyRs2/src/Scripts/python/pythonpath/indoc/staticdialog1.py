#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper, json  # import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from com.sun.star.accessibility import AccessibleRole  # 定数
from com.sun.star.awt import XActionListener, XMenuListener, XMouseListener, XWindowListener, XTextListener, XItemListener
from com.sun.star.awt import MessageBoxButtons, MessageBoxResults, MouseButton, PopupMenuDirection, PosSize, ScrollBarOrientation  # 定数
from com.sun.star.awt import Point, Rectangle, Selection  # Struct
from com.sun.star.awt.MessageBoxType import QUERYBOX  # enum
from com.sun.star.awt.grid import XGridSelectionListener
from com.sun.star.beans import NamedValue  # Struct
from com.sun.star.frame import XFrameActionListener
from com.sun.star.frame.FrameAction import FRAME_UI_DEACTIVATING  # enum
from com.sun.star.i18n.TransliterationModulesNew import FULLWIDTH_HALFWIDTH  # enum
from com.sun.star.util import XCloseListener
from com.sun.star.util import MeasureUnit  # 定数
from com.sun.star.view.SelectionType import MULTI  # enum 
from com.sun.star.lang import Locale  # Struct
from com.sun.star.awt.MenuItemStyle import AUTOCHECK, CHECKABLE
SHEETNAME = "config"  # データを保存するシート名。
def createDialog(xscriptcontext, enhancedmouseevent, dialogtitle, defaultrows=None):  # dialogtitleはダイアログのデータ保存名に使うのでユニークでないといけない。defaultrowsはグリッドコントロールのデフォルトデータ。
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	doc = xscriptcontext.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
	dialogpoint = getDialogPoint(doc, enhancedmouseevent)  # クリックした位置のメニューバーの高さ分下の位置を取得。単位ピクセル。一部しか表示されていないセルのときはNoneが返る。
	if dialogpoint:  # クリックした位置が取得出来た時。
		docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
		containerwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
		toolkit = containerwindow.getToolkit()  # ピアからツールキットを取得。  		
		m = 6  # コントロール間の間隔。
		gridprops = {"PositionX": m, "PositionY": m, "Height": 50, "Width": 100, "ShowRowHeader": False, "ShowColumnHeader": False, "SelectionModel": MULTI}  # グリッドコントロールのプロパティ。
		controlcontainerprops = {"PositionX": 0, "PositionY": 0, "Width": XWidth(gridprops, m), "Height": YHeight(gridprops, m), "BackgroundColor": 0xF0F0F0}  # コントロールコンテナの基本プロパティ。幅は右端のコントロールから取得。高さはコントロール追加後に最後に設定し直す。		
		maTopx = createConverters(containerwindow)  # ma単位をピクセルに変換する関数を取得。
		controlcontainer, addControl = controlcontainerMaCreator(ctx, smgr, maTopx, controlcontainerprops)  # コントロールコンテナの作成。		
		mouselistener = MouseListener(xscriptcontext)
		datarows = getSavedData(doc, "GridDatarows_{}".format(dialogtitle))  # グリッドコントロールの行をconfigシートのragenameから取得する。	
		if datarows is None and defaultrows is not None:  # 履歴がなくデフォルトdatarowsがあるときデフォルトデータを使用。
			datarows = [i if isinstance(i, (list, tuple)) else (i,) for i in defaultrows]  # defaultrowsの要素をリストかタプルでなければタプルに変換する。		
		menulistener = MenuListener(controlcontainer, datarows)  # コンテクストメニューにつけるリスナー。
		items = ("追記", CHECKABLE+AUTOCHECK, {"checkItem": False}),\
				("サイズ保存", CHECKABLE+AUTOCHECK, {"checkItem": True}),\
				(),\
				("新規行挿入", 0, {"setCommand": "insert"}),\
				("行編集", 0, {"setCommand": "edit"}),\
				(),\
		 		("切り取り", 0, {"setCommand": "cut"}),\
				("コピー", 0, {"setCommand": "copy"}),\
				("上に貼り付け", 0, {"setCommand": "pasteabove"}),\
				("下に貼り付け", 0, {"setCommand": "pastebelow"}),\
				(),\
				("選択行を削除", 0, {"setCommand": "delete"}),\
				("全行を削除", 0, {"setCommand": "deleteall"})  # グリッドコントロールにつける右クリックメニュー。
		mouselistener.gridpopupmenu = menuCreator(ctx, smgr)("PopupMenu", items, {"addMenuListener": menulistener})  # 右クリックでまず呼び出すポップアップメニュー。 
		gridcontrol1 = addControl("Grid", gridprops, {"addMouseListener": mouselistener})  # グリッドコントロールの取得。gridは他のコントロールの設定に使うのでコピーを渡す。
		gridmodel = gridcontrol1.getModel()  # グリッドコントロールモデルの取得。
		gridcolumn = gridmodel.getPropertyValue("ColumnModel")  # DefaultGridColumnModel
		column0 = gridcolumn.createColumn()  # 列の作成。
		gridcolumn.addColumn(column0)  # 列を追加。
		griddatamodel = gridmodel.getPropertyValue("GridDataModel")  # GridDataModel
		if datarows:  # 行のリストが取得出来た時。
			griddatamodel.addRows(("",)*len(datarows), datarows)  # グリッドに行を追加。
		rectangle = controlcontainer.getPosSize()  # コントロールコンテナのRectangle Structを取得。px単位。
		rectangle.X, rectangle.Y = dialogpoint  # クリックした位置を取得。ウィンドウタイトルを含めない座標。
		taskcreator = smgr.createInstanceWithContext('com.sun.star.frame.TaskCreator', ctx)
		args = NamedValue("PosSize", rectangle), NamedValue("FrameName", "controldialog")  # , NamedValue("MakeVisible", True)  # TaskCreatorで作成するフレームのコンテナウィンドウのプロパティ。
		dialogframe = taskcreator.createInstanceWithArguments(args)  # コンテナウィンドウ付きの新しいフレームの取得。
		dialogwindow = dialogframe.getContainerWindow()  # ダイアログのコンテナウィンドウを取得。
		dialogframe.setTitle(dialogtitle)  # フレームのタイトルを設定。
		docframe.getFrames().append(dialogframe) # 新しく作ったフレームを既存のフレームの階層に追加する。		
		controlcontainer.createPeer(toolkit, dialogwindow) # ウィンドウにコントロールを描画。 
		frameactionlistener = FrameActionListener()  # FrameActionListener。フレームがアクティブでなくなった時に閉じるため。
		dialogframe.addFrameActionListener(frameactionlistener)  # FrameActionListenerをダイアログフレームに追加。
		controlcontainer.setVisible(True)  # コントロールの表示。
		dialogwindow.setVisible(True) # ウィンドウの表示	
		windowlistener = WindowListener(controlcontainer)
		dialogwindow.addWindowListener(windowlistener) # setVisible(True)でも呼び出されるので、その後でリスナーを追加する。		
		args = doc, controlcontainer, dialogwindow, windowlistener, mouselistener, menulistener
		dialogframe.addCloseListener(CloseListener(args))  # CloseListener。ノンモダルダイアログのリスナー削除用。	
		scrollDown(gridcontrol1)
class CloseListener(unohelper.Base, XCloseListener):  # ノンモダルダイアログのリスナー削除用。
	def __init__(self, args):
		self.args = args
	def queryClosing(self, eventobject, getsownership):  # ノンモダルダイアログを閉じる時に発火。
		dialogframe = eventobject.Source
		doc, controlcontainer, dialogwindow, windowlistener, mouselistener, menulistener = self.args

		
		
		size = controlcontainer.getSize()
		dialogstate = {"append": 0,\
					"Width": size.Width,\
					"Height": size.Height}  # チェックボックスの状態と大きさを取得。
		dialogtitle = dialogframe.getTitle()
		saveData(doc, "dialogstate_{}".format(dialogtitle), dialogstate)
		gridcontrol1 = controlcontainer.getControl("Grid1")
		saveData(doc, "GridDatarows_{}".format(dialogtitle), DATAROWS)
		mouselistener.gridpopupmenu.removeMenuListener(menulistener)
		gridcontrol1.removeMouseListener(mouselistener)
		dialogwindow.removeWindowListener(windowlistener)
		eventobject.Source.removeCloseListener(self)
	def notifyClosing(self, eventobject):
		pass
	def disposing(self, eventobject):  
		pass
def saveData(doc, rangename, obj):	# configシートの名前rangenameにobjをJSONにして保存する。グローバル変数SHEETNAMEを使用。
	namedranges = doc.getPropertyValue("NamedRanges")  # ドキュメントのNamedRangesを取得。
	if rangename in namedranges:  # 名前がある時。
		referredcells = namedranges[rangename].getReferredCells()  # 名前の参照セル範囲を取得。
		if referredcells:  # 参照セル範囲がある時。
			referredcells.setString(json.dumps(obj,  ensure_ascii=False))  # rangenameという名前のセルに文字列でPythonオブジェクトを出力する。
			return
		else:  # 名前があるが参照セル範囲がない時。
			namedranges.removeByName(rangename)  # 名前は重複しているとエラーになるので削除する。	
	if not rangename in namedranges:  # 名前がない時。
		sheets = doc.getSheets()  # シートコレクションを取得。
		if not SHEETNAME in sheets:  # 保存シートがない時。
			sheets.insertNewByName(SHEETNAME, len(sheets))   # 履歴シートを挿入。同名のシートがあるとRuntimeExceptionがでる。
		sheet = sheets[SHEETNAME]  # 保存シートを取得。
		sheet.setPropertyValue("IsVisible", False)  # 非表示シートにする。
		emptyranges = sheet[:, 0].queryEmptyCells()  # 1列目の空セル範囲コレクションを取得。
		if len(emptyranges):  # セル範囲コレクションが取得出来た時。
			cellcursor = sheet.createCursorByRange(emptyranges[0][0, 0])  # 最初のセル範囲の最初のセルからセルカーサーを取得。
			cellcursor.collapseToSize(2, 1)  # 行数1、列数2に変更。(列、行)で指定。
			cellcursor.setDataArray(((rangename, json.dumps(obj, ensure_ascii=False)),))  # セル範囲名を1列目、JSONデータを2列目に代入。
			namedranges.addNewByName(rangename, cellcursor[0, 1].getPropertyValue("AbsoluteName"), cellcursor[0, 1].getCellAddress(), 0)  # 2列目のセルに名前を付ける。名前、式(相対アドレス)、原点となるセル、NamedRangeFlag
def getSavedData(doc, rangename):  # configシートのragenameからデータを取得する。	
	namedranges = doc.getPropertyValue("NamedRanges")  # ドキュメントのNamedRangesを取得。	
	if rangename in namedranges:  # 名前がある時。
		referredcells = namedranges[rangename].getReferredCells()  # 名前が参照しているセル範囲を取得。参照アドレスがエラーのときはNoneが返る。
		if referredcells:
			txt = referredcells.getString()  # 名前が参照しているセルから文字列を取得。
			if txt:
				try:
					return json.loads(txt)  # pyunoオブジェクトは変換できない。
				except json.JSONDecodeError:
					import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
	return None  # 保存された行が取得できない時はNoneを返す。
class MouseListener(unohelper.Base, XMouseListener):  
	def __init__(self, xscriptcontext): 	
		self.xscriptcontext = xscriptcontext
	def mousePressed(self, mouseevent):  # グリッドコントロールをクリックした時。コントロールモデルにはNameプロパティはない。
		gridcontrol = mouseevent.Source  # グリッドコントロールを取得。
		if mouseevent.Buttons==MouseButton.LEFT and mouseevent.ClickCount==2:  # ダブルクリックの時。
			doc = self.xscriptcontext.getDocument()
			selection = doc.getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
			if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択オブジェクトがセルの時。
				griddata = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。
				rowdata = griddata.getRowData(gridcontrol.getCurrentRow())  # グリッドコントロールで選択している行のすべての列をタプルで取得。
				selection.setString(rowdata[0])  # グリッドコントロールは1列と決めつける。
				controller = doc.getCurrentController()  # 現在のコントローラを取得。			
				sheet = controller.getActiveSheet()
				celladdress = selection.getCellAddress()
				nextcell = sheet[celladdress.Row+1, celladdress.Column]  # 下のセルを取得。
				controller.select(nextcell)  # 下のセルを選択。
				nexttxt = nextcell.getString()  # 下のセルの文字列を取得。
				edit1 = gridcontrol.getContext().getControl("Edit1")  # テキストボックスコントロールを取得。				
				edit1.setText(nexttxt)  # テキストボックスコントロールにセルの内容を取得。				
		elif mouseevent.PopupTrigger:  # 右クリックの時。
			rowindex = gridcontrol.getRowAtPoint(mouseevent.X, mouseevent.Y)  # クリックした位置の行インデックスを取得。該当行がない時は-1が返ってくる。
			if rowindex>-1:  # クリックした位置に行が存在する時。
				if not gridcontrol.isRowSelected(rowindex):  # クリックした位置の行が選択状態でない時。
					gridcontrol.deselectAllRows()  # 行の選択状態をすべて解除する。
					gridcontrol.selectRow(rowindex)  # 右クリックしたところの行を選択する。		
				pos = Rectangle(mouseevent.X, mouseevent.Y, 0, 0)  # ポップアップメニューを表示させる起点。
				self.gridpopupmenu.execute(gridcontrol.getPeer(), pos, PopupMenuDirection.EXECUTE_DEFAULT)  # ポップアップメニューを表示させる。引数は親ピア、位置、方向					
	def mouseReleased(self, mouseevent):
		pass
	def mouseEntered(self, mouseevent):
		pass
	def mouseExited(self, mouseevent):
		pass
	def disposing(self, eventobject):
		pass
class MenuListener(unohelper.Base, XMenuListener):
	def __init__(self, controlcontainer, datarows):
		self.controlcontainer = controlcontainer
		self.datarows = datarows
	def itemHighlighted(self, menuevent):
		pass
	def itemSelected(self, menuevent):  # PopupMenuの項目がクリックされた時。どこのコントロールのメニューかを知る方法はない。
		cmd = menuevent.Source.getCommand(menuevent.MenuId)
		controlcontainer = self.controlcontainer
		datarows = DATAROWS
		peer = controlcontainer.getPeer()  # ピアを取得。	
		gridcontrol = controlcontainer.getControl("Grid1")  # グリッドコントロールを取得。
		griddatamodel = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。					
		if cmd=="delete":  # 選択行を削除する。  
			msg = "選択行を削除しますか?"
			msgbox = peer.getToolkit().createMessageBox(peer, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO, SHEETNAME, msg)
			if msgbox.execute()==MessageBoxResults.YES:		
				selectedrows = gridcontrol.getSelectedRows() if griddatamodel.RowCount>1 else (0,)  # 選択行インデックスのタプルを取得。1行だけの時はgetSelectedRows()で取得できない。
				for i in reversed(selectedrows):  # 選択した行インデックスを後ろから取得。逐次検索のときはグリッドコントロールとDATAROWSが一致しないので別に処理する。
					datarows.remove(griddatamodel.getRowData(i))  # 削除するデータ行と一致する要素をDATAROWSから削除。
					griddatamodel.removeRow(i)  # グリッドコントロールから選択行を削除。
		elif cmd=="deleteall":  # 全行を削除する。  	
			msg = "すべての行を削除しますか?"
			msgbox = peer.getToolkit().createMessageBox(peer, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO, SHEETNAME, msg)
			if msgbox.execute()==MessageBoxResults.YES:		
				msg = "本当にすべての行を削除しますか？\n削除したデータは取り戻せません。"
				msgbox = peer.getToolkit().createMessageBox(peer, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO, SHEETNAME, msg)				
				if msgbox.execute()==MessageBoxResults.YES:				
					griddatamodel.removeAllRows()  # グリッドコントロールの行を全削除。
					datarows.clear()  # 全データ行をクリア。	
		global DATAROWS
		DATAROWS = datarows			
	def itemActivated(self, menuevent):
		pass
	def itemDeactivated(self, menuevent):
		pass   
	def disposing(self, eventobject):
		pass
class FrameActionListener(unohelper.Base, XFrameActionListener):
	def frameAction(self, frameactionevent):
		if frameactionevent.Action==FRAME_UI_DEACTIVATING:  # フレームがアクティブでなくなった時。TopWindowListenerのwindowDeactivated()だとウィンドウタイトルバーをクリックしただけで発火してしまう。
			frameactionevent.Frame.removeFrameActionListener(self)  # フレームにつけたリスナーを除去。
			frameactionevent.Frame.close(True)
	def disposing(self, eventobject):
		pass
def resizeControls(controlcontainer, oldwidth, oldheight, newwidth, newheight):	 # ウィンドウの大きさの変更に合わせてコントロールの位置と大きさを変更。ウィンドウの大きさはここで変更するとコントロールが正しく移動できない。
	gridcontrol1 = controlcontainer.getControl("Grid1")
	gridcontrol1rect = gridcontrol1.getPosSize()
	m = gridcontrol1rect.X  # コントロール間の間隔をチェックボックスのPositionXから取得。
	minwidth = minheight = m*3  # 幅下限、高さ下限、を取得。
	if newwidth<minwidth:  # 変更後のコントロールコンテナの幅を取得。サイズ下限より小さい時は下限値とする。
		newwidth = minwidth
	if newheight<minheight:  # 変更後のコントロールコンテナの高さを取得。サイズ下限より小さい時は下限値とする。
		newheight = minheight
	diff_width = newwidth - oldwidth  # 幅変化分
	diff_height = newheight - oldheight  # 高さ変化分		
	applyDiff = createApplyDiff(diff_width, diff_height)  # コントロールの位置と大きさを変更する関数を取得。
	controlcontainer.setPosSize(0, 0, newwidth, newheight, PosSize.SIZE)  # コントロールコンテナの大きさを変更する。
	applyDiff(gridcontrol1, PosSize.SIZE)	
	return newwidth, newheight  # 下限制限後のウィンドウサイズを返す。
class WindowListener(unohelper.Base, XWindowListener):
	def __init__(self, controlcontainer):
		size = controlcontainer.getSize()
		self.oldwidth = size.Width  # 次の変更前の幅として取得。
		self.oldheight = size.Height  # 次の変更前の高さとして取得。		
		self.controlcontainer = controlcontainer
	def windowResized(self, windowevent):
		newwidth, newheight = resizeControls(self.controlcontainer, self.oldwidth, self.oldheight, windowevent.Width, windowevent.Height)  # 変化分で計算する。コントロールが表示されないほど小さくされると次から表示がおかしくなる。
		self.oldwidth = newwidth  # 次の変更前の幅として取得。
		self.oldheight = newheight  # 次の変更前の高さとして取得。
		scrollDown(self.controlcontainer.getControl("Grid1"))
	def windowMoved(self, windowevent):
		pass
	def windowShown(self, eventobject):
		pass
	def windowHidden(self, eventobject):
		pass
	def disposing(self, eventobject):
		pass
def createApplyDiff(diff_width, diff_height):		
	def applyDiff(control, possize):  # 第2引数でウィンドウサイズの変化分のみ適用するPosSizeを指定。
		rectangle = control.getPosSize()  # 変更前のコントロールの位置大きさを取得。
		control.setPosSize(rectangle.X+diff_width, rectangle.Y+diff_height, rectangle.Width+diff_width, rectangle.Height+diff_height, possize)  # Flagsで変更する値のみ指定。変更しない値は0(でもなんでもよいはず)。
	return applyDiff	
def scrollDown(gridcontrol):  # グリッドコントロールを下までスクロールする。		
	accessiblecontext = gridcontrol.getAccessibleContext()  # グリッドコントロールのAccessibleContextを取得。
	for i in range(accessiblecontext.getAccessibleChildCount()):  # 子要素をのインデックスを走査する。
		child = accessiblecontext.getAccessibleChild(i)  # 子要素を取得。
		if child.getAccessibleContext().getAccessibleRole()==AccessibleRole.SCROLL_BAR:  # スクロールバーの時。
			if child.getOrientation()==ScrollBarOrientation.VERTICAL:  # 縦スクロールバーの時。
				child.setValue(child.getMaximum())  # 最大値にスクロールさせる。
				break				
def XWidth(props, m=0):  # 左隣のコントロールからPositionXを取得。mは間隔。
	return props["PositionX"] + props["Width"] + m  	
def YHeight(props, m=0):  # 上隣のコントロールからPositionYを取得。mは間隔。
	return props["PositionY"] + props["Height"] + m
def getDialogPoint(doc, enhancedmouseevent):  # クリックした位置x yのタプルで返す。但し、一部しか見えてないセルの場合はNoneが返る。TaskCreatorのRectangleには画面の左角からの座標を渡すが、ウィンドウタイトルバーは含まれない。
	controller = doc.getCurrentController()  # 現在のコントローラを取得。
	docframe = controller.getFrame()  # フレームを取得。
	containerwindow = docframe.getContainerWindow()  # コンテナウィドウの取得。
	accessiblecontextparent = containerwindow.getAccessibleContext().getAccessibleParent()  # コンテナウィンドウの親AccessibleContextを取得する。フレームの子AccessibleContextになる。
	accessiblecontext = accessiblecontextparent.getAccessibleContext()  # AccessibleContextを取得。
	for i in range(accessiblecontext.getAccessibleChildCount()): 
		childaccessiblecontext = accessiblecontext.getAccessibleChild(i).getAccessibleContext()
		if childaccessiblecontext.getAccessibleRole()==49:  # ROOT_PANEの時。
			rootpanebounds = childaccessiblecontext.getBounds()  # Yアトリビュートがウィンドウタイトルバーの高さになる。
			break 
	else:
		return  # ウィンドウタイトルバーのAccessibleContextが取得できなかった時はNoneを返す。
	componentwindow = docframe.getComponentWindow()  # コンポーネントウィンドウを取得。
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
								x = sourcepointonscreen.X + enhancedmouseevent.X  # クリックした位置の画面の左上角からのXの取得。
								y = sourcepointonscreen.Y + enhancedmouseevent.Y + rootpanebounds.Y  # クリックした位置からメニューバーの高さ分下の位置の画面の左上角からのYの取得									
								return x, y
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
def controlcontainerMaCreator(ctx, smgr, maTopx, containerprops):  # ma単位でコントロールコンテナと、それにコントロールを追加する関数を返す。まずコントロールコンテナモデルのプロパティを取得。UnoControlDialogElementサービスのプロパティは使えない。propsのキーにPosSize、値にPOSSIZEが必要。
	container = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainer", ctx)  # コントロールコンテナの生成。
	container.setPosSize(*maTopx(containerprops.pop("PositionX"), containerprops.pop("PositionY")), *maTopx(containerprops.pop("Width"), containerprops.pop("Height")), PosSize.POSSIZE)
	containermodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainerModel", ctx)  # コンテナモデルの生成。
	containermodel.setPropertyValues(tuple(containerprops.keys()), tuple(containerprops.values()))  # コンテナモデルのプロパティを設定。存在しないプロパティに設定してもエラーはでない。
	container.setModel(containermodel)  # コンテナにコンテナモデルを設定。
	container.setVisible(False)  # 描画中のものを表示しない。
	def addControl(controltype, props, attrs=None):  # props: コントロールモデルのプロパティ、キーにPosSize、値にPOSSIZEが必要。attr: コントロールの属性。
		name = props.pop("Name") if "Name" in props else _generateSequentialName(controltype) # サービスマネージャーからインスタンス化したコントロールはNameプロパティがないので、コントロールコンテナのaddControl()で名前を使うのみ。
		controlidl = "com.sun.star.awt.grid.UnoControl{}".format(controltype) if controltype=="Grid" else "com.sun.star.awt.UnoControl{}".format(controltype)  # グリッドコントロールだけモジュールが異なる。
		control = smgr.createInstanceWithContext(controlidl, ctx)  # コントロールを生成。
		control.setPosSize(*maTopx(props.pop("PositionX"), props.pop("PositionY")), *maTopx(props.pop("Width"), props.pop("Height")), PosSize.POSSIZE)  # ピクセルで指定するために位置座標と大きさだけコントロールで設定。
		controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
		control.setModel(controlmodel)  # コントロールにコントロールモデルを設定。
		container.addControl(name, control)  # コントロールをコントロールコンテナに追加。
		if attrs is not None:  # Dialogに追加したあとでないと各コントロールへの属性は追加できない。
			for key, val in attrs.items():  # メソッドの引数がないときはvalをNoneにしている。
				if val is None:
					getattr(control, key)()
				else:
					getattr(control, key)(val)
		return control  # 追加したコントロールを返す。
	def _createControlModel(controltype, props):  # コントロールモデルの生成。
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
