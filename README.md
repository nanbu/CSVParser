# CSVParser

 * VBAでCSVファイル(Comma-Separated Values形式のテキストファイル)を読み取るクラスです。
 * Windows、MacどちらのVBAでも動作します。

 ## 特徴
 
 * RFC4180の仕様を満たしています。よって、フィールド内にカンマ、改行、ダブルクオートを含む場合でも正しく読み取ることができます。
 * Windows、MacいずれのVBAでも動作します。
 * 読み込めるCSVファイルのレコード数に制限はありません。
 * クラスとして実装しているので他のモジュールに干渉することを気にせず導入できます。ファイルオープン、クローズ
 * CSVファイルを1レコードずつ読み取ります。
 * 1レコードはCollectionオブジェクトとして取得します。
 * 各フィールドは文字列で取得します。
 * 空行は「空文字フィールドが1つ存在するレコード」として読み取ります。

## 制限事項
CSVファイルの文字コードは、VBAのInput関数で読み取れるものと同じになります。（文字コードは指定できません。）日本語版のExcelではShift_JISになるようです。

## インストール

Visual Basic EditorよりCSVParser.clsをインポートしてください。

## 使い方

 1. CSVParserをインスタンス化します。
 2. OpenFile(Filename)メソッドでCSVファイルを開きます。
 3. EndOfDataプロパティでデータに読み取り可能なレコードがあるかを確認します。
 4. ReadFieldsメソッドで1レコード分の各フィールドを取り出します。
 5. CloseFileメソッドでCSVファイルを閉じます。

## API

	OpenFile(Filename)
Filename: CSVファイルのパス  
CSVファイルを開きます。

	EndOfData() As Boolean
データに読み取り可能なレコードがあるかを確認します。読み取り専用のプロパティです。  
読み取り可能なデータが残っていない場合はTrue、それ以外の場合はFalseになっています。  
ファイルを開いていない場合はTrueになります。

	ReadFields() As Collection
1レコードを読み取り、各フィールドをCollectionオブジェクトに格納された文字列として返します。  
このメソッドを実行すると読み取り対象レコードが1つ次に進みます。  
読み取りが完了した場合には同時にCSVファイルを閉じます。

	CloseFile()
OpenFile(Filename)メソッドで開いたCSVファイルを閉じます。
CSVファイルが開いていない場合は何もしません。

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
