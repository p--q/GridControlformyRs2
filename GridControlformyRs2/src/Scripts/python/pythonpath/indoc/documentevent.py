#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# ドキュメントイベントについて。
def documentOnLoad(xscriptcontext):  # ドキュメントを開いた時。リスナー追加前。
	doc = xscriptcontext.getDocument()  
	namedranges = doc.getPropertyValue("NamedRanges")  # ドキュメントのNamedRangesを取得。
	[namedranges.removeByName(i.getName()) for i in namedranges if not i.getReferredCells()]  # 参照範囲がエラーの名前を削除する。	
def documentUnLoad(xscriptcontext):  # ドキュメントを閉じた時。リスナー削除後。
	pass
