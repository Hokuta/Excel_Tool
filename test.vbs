
'
'引数リスト
' FileName:固定長テキストファイルのフルパス
' QuotMark:文字列の引用符。省略すると引用記号なし。
' Addition:csvは追加モードか？追加モードならTrue
' FieldSize:固定長ファイルのフォーマット(各フィールドの長さ)をカンマ(,)区切りで指定。
'            指定数は任意。
'実行結果:固定長テキストファイルと同じディレクトリに
'		 同名のcsvファイルが作られる。
'   -----------------------------------------------------
Function csvConv(ByVal FileName As String, ByVal QuotMark As Variant, ByVal Addition As Boolean, ParamArray ByVal FieldSize() As Variant)

	Dim FSO As Object
	Dim Stream(1) As Object
	Dim txt As String
	Dim i As Integer, n As Integer

	Dim strD(1) As String
	QuotMark = Nz(QuotMark, "")
	Addition = 2 + Abs(Addition) * 6
	strD(0) = QuotMark
	strD(1) = strD(0) & "," & strD(0)

	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set Stream(0) = FSO.OpenTextFile(FileName, 1)
	Set Stream(1) = FSO.OpenTextFile(Replace(FileName, FSO.GetExtensionName(FileName), "csv"), Addition, True)
  Do Until Stream(0).AtEndOfStream
	 n = 0
	 txt = Stream(0).ReadLine
	   For i = 0 To UBound(FieldSize)
		 txt = Left$(txt, n + i + FieldSize(i)) & _
			   strD(1) & _
			   Mid$(txt, n + i + FieldSize(i) + 1)
		 n = n + FieldSize(i)
	   Next i
	 txt = strD(0) & txt & strD(0)
	 Stream(1).WriteLine txt
   Loop
   Stream(0).Close: Stream(1).Close
   Set FSO = Nothing
 End Function
