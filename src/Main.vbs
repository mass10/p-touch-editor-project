'==============================================================================
' アプリケーションのエントリーポイント
'==============================================================================
Sub Main()

	'b-PACオブジェクトを生成
	Set ObjDoc = CreateObject("bpac.Document")

	'P-touch Editorで作成したテンプレートファイルをオープン
	'lbxファイルをVBSファイルと同じフォルダに置いておく
	bRet = ObjDoc.Open("..\DATA\template.lbx")

	If bRet Then '正常にオープン？

		'住所
		ObjDoc.GetObject("AddressText").Text = "999-9999 東京都千代田区中央 1-1-1 テストビル"

		'おなまえ
		ObjDoc.GetObject("NameText").Text = "Jimi Hendrix 様"

		'印刷を実行
		ObjDoc.StartPrint "DocumentName", bpoAutoCut
		ObjDoc.PrintOut 1, bpoAutoCut
		ObjDoc.EndPrint

		'印刷を実行(2)
		' ObjDoc.StartPrint "DocumentName", bpoAutoCut
		' ObjDoc.PrintOut 1, bpoAutoCut
		' ObjDoc.EndPrint

		'印刷を実行(3)
		' ObjDoc.StartPrint "DocumentName", bpoAutoCut
		' ObjDoc.PrintOut 1, bpoAutoCut
		' ObjDoc.EndPrint

		ObjDoc.Close

	End If

	'b-PACオブジェクトを解放
	Set ObjDoc = Nothing

End Sub

Call Main()
