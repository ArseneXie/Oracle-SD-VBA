Attribute VB_Name = "AutoSDTable"
Sub CreateDataDic()

'
' CreateDataDic Macro in Word
' @authur arsene@readycom.com.tw
' @version 1.3  using clipboard
'
'
    Dim curTable As Table
    Dim r As Integer
    Dim i As Integer
    Dim j As Integer
    Dim p As Integer
    Dim x As Integer
    Dim sqlScript As String
    Dim readLine() As String
    Dim dicDetail() As String
    Dim tableName As String
    Dim columnName As String
    Dim colType As String
    Dim colTypeLen As String
    Dim colCount As Integer
    Dim commentStr As String
    Dim ownerStr As String
    Dim myData As DataObject
   
    If Selection.Information(wdWithInTable) = True Then
        With Selection
        Set curTable = .Tables(1)
            Set myData = New DataObject
            myData.GetFromClipboard
            sqlScript = myData.GetText
            readLine = Split(sqlScript, Chr(10))
            
            MsgBox sqlScript
            'MsgBox readLine(0)
            
            i = 0
            Do While i <= UBound(readLine)
                readLine(i) = Trim(Replace(readLine(i), Chr(13), ""))
                i = i + 1
            Loop
            
            x = InStr(1, readLine(0), ".", 1)
            If x > 0 Then
                ownerStr = Trim(Replace(UCase(Left(readLine(0), x)), "CREATE TABLE", ""))
            Else
                ownerStr = ""
            End If
            'MsgBox ownerStr
            tableName = Trim(Replace(Replace(UCase(readLine(0)), "CREATE TABLE", ""), ownerStr, ""))
            curTable.Cell(1, 1).Range.Select
            Selection.TypeText Text:="Table Name: " + tableName
            i = 0
            colCount = 0
            ReDim dicDetail(1 To UBound(readLine), 1 To 4)
            ' Hanle main script
            Do While i <= UBound(readLine)
                If Left(readLine(i), 1) = "(" Then
                    colCount = i
                    r = 0
                ElseIf Left(readLine(i), 1) = ")" Then
                    colCount = i - colCount - 1
                    'MsgBox colCount
                    i = UBound(readLine)
                ElseIf colCount > 0 Then
                    r = r + 1
                    p = InStr(1, readLine(i), Chr(32), 1)
                    
                    columnName = UCase(Left(readLine(i), p - 1))
                    colType = Replace(Trim(UCase(Replace(Right(readLine(i), Len(readLine(i)) - p), ",", ""))), "NOT NULL", "")
                    
                    x = InStr(1, colType, "(", 1)
                    If x > 0 Then
                        colTypeLen = Replace(Right(colType, Len(colType) - x), ")", "")
                        colType = Left(colType, x - 1)
                    Else
                        colTypeLen = ""
                    End If
                    dicDetail(r, 1) = columnName
                    dicDetail(r, 2) = colType
                    dicDetail(r, 3) = colTypeLen
                End If
                i = i + 1
            Loop
            
        
            'Handle comment script
            i = colCount + 3
            
            Do While i <= UBound(readLine)
                If InStr(1, readLine(i), "comment on column", 1) > 0 Then
                    x = InStr(1, readLine(i), ownerStr + tableName + ".", 1)
                    x = x + Len(tableName) + Len(ownerStr)
                    columnName = UCase(Right(readLine(i), Len(readLine(i)) - x))
                    'MsgBox columnName
                    For j = 1 To colCount
                        If dicDetail(j, 1) = columnName Then
                               p = j
                               Exit For
                        End If
                    Next
                ElseIf InStr(1, UCase(readLine(i)), "IS", 1) > 0 Then
                    x = InStr(1, readLine(i), "'", 1)
                    commentStr = Right(readLine(i), Len(readLine(i)) - x)
                    commentStr = Left(commentStr, Len(commentStr) - 2)
                    'MsgBox commentStr
                    dicDetail(j, 4) = commentStr
                End If
                i = i + 1
            Loop
            
            'Write out
            r = 1
            Do While r <= colCount
                x = r + 2
                curTable.Cell(x, 2).Range.Select
                Selection.TypeText Text:=dicDetail(r, 1)
                curTable.Cell(x, 3).Range.Select
                Selection.TypeText Text:=dicDetail(r, 2)
                curTable.Cell(x, 4).Range.Select
                Selection.TypeText Text:=dicDetail(r, 3)
                curTable.Cell(x, 5).Range.Select
                Selection.TypeText Text:=dicDetail(r, 4)
                If r < colCount Then
                    curTable.Cell(x, 5).Range.Select
                    Selection.MoveRight Unit:=wdCell
                End If
                r = r + 1
            Loop
            
            
        End With
    End If
End Sub


Sub CreateTable()
'
' CreateTable Macro in Word
' @authur arsene@readycom.com.tw
' @version 1.3 using clipboard
'
'
    Dim curTable As Table
    Dim r As Integer
    Dim x As Integer
    Dim tableScript As String
    Dim commentScript As String
    Dim tableName As String
    Dim columnName As String
    Dim colType As String
    Dim commentStr As String
    Dim iChar As Integer
    Dim rstStr As String
    Dim myData As DataObject
    tableScript = "CREATE TABLE "
    If Selection.Information(wdWithInTable) = True Then
        With Selection
        Set curTable = .Tables(1)
            curTable.Cell(1, 1).Range.Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            tableName = Trim(Replace(Replace(Replace(Selection.Text, "Table Name", ""), ":", ""), "¡G", ""))
            tableScript = tableScript + tableName + Chr(10) + "(" + Chr(10)
            r = 3
            Do While r <= curTable.Rows.Count
                curTable.Cell(r, 2).Range.Select
                Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                columnName = Selection.Text
                tableScript = tableScript + columnName + " "
                curTable.Cell(r, 3).Range.Select
                Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                colType = Selection.Text
                If UCase(colType) = "VARCHAR2" Then
                    curTable.Cell(r, 4).Range.Select
                    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                    iChar = Asc(Selection.Text)
                    If iChar = 13 Then
                        colType = colType + "(ERROR)"
                    Else
                        colType = colType + "(" + Selection.Text + ")"
                    End If
                End If
                If r < curTable.Rows.Count Then
                    tableScript = tableScript + colType + "," + Chr(10)
                Else
                    tableScript = tableScript + colType + Chr(10) + ");"
                End If
                ' handle commentScript
                curTable.Cell(r, 5).Range.Select
                Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                commentStr = Selection.Text
                x = InStr(Selection.Text, Chr(13))
                If x > 0 Then
                    commentStr = Left(commentStr, x - 1)
                End If
                commentScript = commentScript + Chr(10) + "comment on column " + tableName + "." + columnName + Chr(10) + "  IS '" + commentStr + "';"
                r = r + 1
            Loop
            
            rstStr = tableScript + commentScript
            
            Set myData = New DataObject
            myData.SetText rstStr
            myData.PutInClipboard

            MsgBox "Already generate the script in the clipboard"

        End With
    End If
End Sub
