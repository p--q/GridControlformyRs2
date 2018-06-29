#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
from indoc import documentevent, dialogs  # 相対インポートは不可。
SHEETNAME = "config"  # データを保存するシート名。
def getModule(sheetname):  # シート名に応じてモジュールを振り分ける関数。
	if sheetname is None:  # シート名でNoneが返ってきた時はドキュメントイベントとする。
		return documentevent
	else:  # シート名が渡された時。
		return dialogs
	return None  # モジュールが見つからなかった時はNoneを返す。
