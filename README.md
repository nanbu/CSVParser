# CSVParser

 * VBAでCSVファイルを読み込むクラスです。
 * CSVファイルを1レコードずつ読み込ます。
 * 1レコードはCollectionオブジェクトとして取得します。
 * 各フィールドは文字列で取得します。

## インストール

Visual Basic EditorよりCSVParser.clsをインポートしてください。

## サンプル

	Sub Test
		Dim Filename
		Dim Parser As CSVParser
		Dim Fields As Collection
		Dim Field
		
		Filename = Application.GetOpenFilename
		If VarType(Filename) = vbBoolean Then
			Exit Sub
		End If
		
		Set Parser = New CSVParser
		Parser.OpenFile Filename
		Do Until Parser.EndOfData
			Set Fields = Parser.ReadFields
			For Each Field In Fields
				Debug.Print Field
			Next
		Loop
		Parser.CloseFile
	End Sub
