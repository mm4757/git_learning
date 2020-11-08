Option Explicit

' CSVのファイルを読み込み, アクティブブック内のすべての文字列を置換する.
Public Sub ReplaceByCSV()

    ' 置換前文字列
    Dim beforeReplace() As String
    ' 置換後文字列
    Dim afterReplace() As String

    ' CSVファイルを読み込む
    ' 対象ファイルのパスを取得する.
    Dim path As String
    path = Application.GetOpenFilename("Comma-Separated Values, *.csv")
    
    ' CSVファイルをオープンする.
    Dim nFile As Integer
    ' 使用可能なファイル番号を取得
    nFile = FreeFile
    ' ファイルをファイルナンバーnFileで読専で開く
    Open path For Input As #nFile
    
    ' 読み込んだファイルから置換前と置換後の文字列を抽出
    Dim buf As String
    Dim i As Integer
    i = 0
    Dim tmp As Variant ' 一時配列
    Do Until EOF(nFile)
    
        Line Input #nFile, buf
        
        tmp = Split(buf, ",")
        ' allocationする
        ReDim Preserve beforeReplace(i + 1)
        ReDim Preserve afterReplace(i + 1)
        beforeReplace(i) = tmp(0)
        afterReplace(i) = tmp(1)
        
        i = i + 1
        
    Loop
    
    ' アクティブブック内の全文字列を置換する
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets

        For i = 0 To UBound(beforeReplace)
            ws.UsedRange.Replace What:=beforeReplace(i), Replacement:=afterReplace(i), LookAt:=xlPart
        Next i
        
    Next ws
    
End Sub
