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
	dialogpoint = getDialogPoint(doc, enhancedmouseevent)  # クリックした位置のメニューバーの高さ分下の位置を取得。単位ピクセル。一部しか表示されていないセルのときはNoneが返る。
	if dialogpoint:
		docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
		containerwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
		toolkit = containerwindow.getToolkit()  # ピアからツールキットを取得。  
		m = 6  # コントロール間の間隔。
		h = 12
		gridprops = {"PositionX": m, "PositionY": m, "Width": 104, "Height": 50, "ShowRowHeader": False, "ShowColumnHeader": False, "VScroll": True}  # グリッドコントロールのプロパティ。
		textboxprops = {"PositionX": m, "PositionY": YHeight(gridprops, 2), "Width": gridprops["Width"], "Height": h, "Text": doc.getCurrentSelection().getString()}  # テクストボックスコントロールのプロパティ。
		checkboxprops = {"PositionX": m, "PositionY": YHeight(textboxprops, 4), "Width": 42, "Height": h, "Label": "~サイズ保存", "State": 1}  # チェックボックスコントロールのプロパティ。
		buttonprops = {"PositionY": YHeight(textboxprops, 4), "Width": 30, "Height": h+2, "Label": "Enter"}  # ボタンのプロパティ。PushButtonTypeの値はEnumではエラーになる。VerticalAlignではtextboxと高さが揃わない。
		buttonprops.update(PositionX=XWidth(gridprops)-buttonprops["Width"])
		controlcontainerprops = {"PositionX": 0, "PositionY": 0, "Width": XWidth(gridprops, m), "Height": YHeight(buttonprops, m), "BackgroundColor": 0xF0F0F0}  # コントロールコンテナの基本プロパティ。幅は右端のコントロールから取得。高さはコントロール追加後に最後に設定し直す。		
		sizes = getSavedData(doc, "dialogsize")
		if sizes:
			controlcontainerprops.update(Width=sizes[0], Height=sizes[1])
		maTopx = createConverters(containerwindow)  # ma単位をピクセルに変換する関数を取得。
		controlcontainer, addControl = controlcontainerMaCreator(ctx, smgr, maTopx, controlcontainerprops)  # コントロールコンテナの作成。		
		gridselectionlistener = GridSelectionListener()
		gridcontrol1 = addControl("Grid", gridprops, {"addSelectionListener": gridselectionlistener})  # グリッドコントロールの取得。gridは他のコントロールの設定に使うのでコピーを渡す。
		gridmodel = gridcontrol1.getModel()  # グリッドコントロールモデルの取得。
		gridcolumn = gridmodel.getPropertyValue("ColumnModel")  # DefaultGridColumnModel
		column0 = gridcolumn.createColumn()  # 列の作成。
		gridcolumn.addColumn(column0)  # 列を追加。
		griddata = gridmodel.getPropertyValue("GridDataModel")  # GridDataModel
		datarows = getSavedData(doc, "GridDatarows")  # グリッドコントロールの行をconfigシートのragenameから取得する。	

		now = datetime.now()  # 現在の日時を取得。
		t = now.isoformat()

		if datarows:  # 行のリストが取得出来た時。
			griddata.insertRows(0, ("",)*len(datarows), datarows)  # グリッドに行を挿入。
		else:
			griddata.addRow("", (t,))  # 新規行を追加。
		addControl("Edit", textboxprops)  
		addControl("CheckBox", checkboxprops)  
		actionlistener = ActionListener(xscriptcontext)
		addControl("Button", buttonprops, {"addActionListener": actionlistener, "setActionCommand": "enter"})  
		rectangle = controlcontainer.getPosSize()  # コントロールコンテナのRectangle Structを取得。px単位。
		rectangle.X, rectangle.Y = dialogpoint  # クリックした位置を取得。ウィンドウタイトルを含めない座標。
		taskcreator = smgr.createInstanceWithContext('com.sun.star.frame.TaskCreator', ctx)
		args = NamedValue("PosSize", rectangle), NamedValue("FrameName", "controldialog")  # , NamedValue("MakeVisible", True)  # TaskCreatorで作成するフレームのコンテナウィンドウのプロパティ。
		dialogframe = taskcreator.createInstanceWithArguments(args)  # コンテナウィンドウ付きの新しいフレームの取得。
		dialogwindow = dialogframe.getContainerWindow()  # ダイアログのコンテナウィンドウを取得。
		dialogframe.setTitle("履歴")  # フレームのタイトルを設定。
		docframe.getFrames().append(dialogframe) # 新しく作ったフレームを既存のフレームの階層に追加する。			
		controlcontainer.createPeer(toolkit, dialogwindow) # ウィンドウにコントロールを描画。 
		frameactionlistener = FrameActionListener()  # FrameActionListener。フレームがアクティブでなくなった時に閉じるため。
		dialogframe.addFrameActionListener(frameactionlistener)  # FrameActionListenerをダイアログフレームに追加。
		controlcontainer.setVisible(True)  # コントロールの表示。
		dialogwindow.setVisible(True) # ウィンドウの表示	
		windowlistener = WindowListener(controlcontainer)
		dialogwindow.addWindowListener(windowlistener) # setVisible(True)でも呼び出されるので、その後でリスナーを追加する。		
		args = doc, controlcontainer, gridselectionlistener, actionlistener, dialogwindow, windowlistener
		dialogframe.addCloseListener(CloseListener(args))  # CloseListener。ノンモダルダイアログのリスナー削除用。	
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
		doc, controlcontainer, gridselectionlistener, actionlistener, dialogwindow, windowlistener = self.args
		
		controlcontainer
		
		
		gridcontrol1 = controlcontainer.getControl("Grid1")
		saveGridRows(doc, gridcontrol1)
		gridcontrol1.removeSelectionListener(gridselectionlistener)
		buttoncontrol1 = controlcontainer.getControl("Button1")
		buttoncontrol1.removeActionListener(actionlistener)
		dialogwindow.removeWindowListener(windowlistener)
		eventobject.Source.removeCloseListener(self)
	def notifyClosing(self, eventobject):
		pass
	def disposing(self, eventobject):  
		eventobject.Source.removeCloseListener(self)
class WindowListener(unohelper.Base, XWindowListener):
	def __init__(self, args):
		self.args = args	
		self.oldwidth = 0  # 次の変更前の幅として取得。
		self.oldheight = 0  # 次の変更前の高さとして取得。	
		
		
			
	def windowResized(self, windowevent):  # 変化分で計算する。コントロールが表示されないほど小さくされると次から表示がおかしくなる。
		
		
		
		controlcontainer = self.args
		gridcontrol1 = controlcontainer.getControl("Grid1")
		editcontrol1 = controlcontainer.getControl("Edit1")
		checkboxcontrol1 = controlcontainer.getControl("CheckBox1")
		buttoncontrol1 = controlcontainer.getControl("Button1")
		checkboxrect = checkboxcontrol1.getPosSize()
		m = checkboxrect.X
		minwidth = checkboxrect.Width + buttoncontrol1.getSize().Width + m*2  # 幅下限を取得。
		minheight = checkboxrect.Height*3 + m*4  # 高さ下限を取得。
		newwidth = windowevent.Width if windowevent.Width>minwidth else minwidth  # 変更後のコントロールコンテナの幅を取得。サイズ下限より小さい時は下限値とする。
		newheight = windowevent.Height if windowevent.Height>minheight else minheight  # 変更後のコントロールコンテナの高さを取得。サイズ下限より小さい時は下限値とする。
		diff_width = newwidth - self.oldwidth  # 幅変化分
		diff_height = newheight - self.oldheight  # 高さ変化分		
		applyDiff = self._createApplyDiff(diff_width, diff_height)
		controlcontainer.setPosSize(0, 0, newwidth, newheight, PosSize.SIZE)  # Flagsで変更する値のみ指定。変更しない値は0(でもなんでもよいはず)。
		applyDiff(gridcontrol1, PosSize.SIZE)
		applyDiff(editcontrol1, PosSize.Y+PosSize.WIDTH)
		applyDiff(checkboxcontrol1, PosSize.Y)
		applyDiff(buttoncontrol1, PosSize.POS)
		
		
		
		self.oldwidth = newwidth  # 次の変更前の幅として取得。
		self.oldheight = newheight  # 次の変更前の高さとして取得。
		
		
	def _createApplyDiff(self, diff_width, diff_height):		
		def applyDiff(control, possize):  # 第2引数でウィンドウサイズの変化分のみ適用するPosSizeを指定。
			rectangle = control.getPosSize()  # 変更前のコントロールの位置大きさを取得。
			control.setPosSize(rectangle.X+diff_width, rectangle.Y+diff_height, rectangle.Width+diff_width, rectangle.Height+diff_height, possize)
		return applyDiff
	def windowMoved(self, windowevent):
		pass
	def windowShown(self, eventobject):
		pass
	def windowHidden(self, eventobject):
		pass
	def disposing(self, eventobject):
		eventobject.Source.removeWindowListener(self)
class ActionListener(unohelper.Base, XActionListener):
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
	def actionPerformed(self, actionevent):
		xscriptcontext = self.xscriptcontext
		cmd = actionevent.ActionCommand
		if cmd=="enter":
			edit1 = actionevent.Source.getContext().getControl("Edit1")
			doc = xscriptcontext.getDocument()  
			selection = doc.getCurrentSelection()
			if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択オブジェクトがセルの時。
				selection.setString(edit1.getText())
				#  下のセルに移動。
				#  グリッドにデーターを追加。
				
				
				
	def disposing(self, eventobject):
		eventobject.Source.removeActionListener(self)
def saveGridRows(doc, gridcontrol):  # グリッドコントロールの行をhistoryシートのragenameに保存する。		
	griddatamodel = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModel
	datarows = []  # 重複を避けて行を取得。
	for i in range(griddatamodel.RowCount)[::-1]:  # 下から行を取得。
		datarow = griddatamodel.getRowData(i)  # グリッドコントロールの1行を取得。
		if not datarow in datarows:
			datarows.append(datarow)
	datarows.reverse()  # 使用順に戻す。
	

	
	# 1番下を残して重複データを削除。
	# 上限数設定
	
	saveData(doc, "GridDatarows", datarows)
def saveData(doc, rangename, obj):	# configシートの名前rangenameにobjをJSONにして保存する。
	namedranges = doc.getPropertyValue("NamedRanges")  # ドキュメントのNamedRangesを取得。
	if not rangename in namedranges:  # 名前がない時。名前は重複しているとエラーになる。
		sheets = doc.getSheets()  # シートコレクションを取得。
		sheetname = "config"  # 履歴シート名。
		if not sheetname in sheets:  # 履歴シートがない時。
			sheets.insertNewByName(sheetname, len(sheets))   # 履歴シートを挿入。同名のシートがあるとRuntimeExceptionがでる。
		sheet = sheets[sheetname]  # 履歴シートを取得。
		sheet.setPropertyValue("IsVisible", False)  # 非表示シートにする。
		emptyranges = sheet[:, :2].queryEmptyCells()  # 2列目までの最初の空セル範囲コレクションを取得。
		if len(emptyranges):  # セル範囲コレクションが取得出来た時。
			emptyrange = emptyranges[0]  # 最初のセル範囲を取得。
			emptyrange[0, 0].setString(rangename)  # 1列目に名前を表示する。
			namedranges.addNewByName(rangename, emptyrange[0, 1].getPropertyValue("AbsoluteName"), emptyrange[0, 1].getCellAddress(), 0)  # 2列目のセルに名前を付ける。名前、式(相対アドレス)、原点となるセル、NamedRangeFlag
	namedranges[rangename].getReferredCells().setString(json.dumps(obj,  ensure_ascii=False))  # rangenameという名前のセルに文字列でリストを出力する。
def getSavedData(doc, rangename):  # configシートのragenameからデータを取得する。	
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
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
