Attribute VB_Name = "Module1"
Option Explicit

'Convert to Windows-1251 if inserted into Excel!

'---------------------------------------------------------------------
'Copyright 2017-2018 Дмитрий Евдокимов
'    http://dievdo.ru
'
'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'
'    http://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.
'---------------------------------------------------------------------

Const F440Pin = "D:\OD\FORMS\F440p\in\"
Const F440Prep = "D:\OD\FORMS\F440p\rep\"
Const DatePath = "yyyy\\MM\\dd\\"

'---------------------------------------------------------------------

Const ColID = 1
Const ColDate = 2
Const ColTime = 3
Const ColType = 4
Const ColMZFile = 5
Const ColRepDate = 6
Const ColRepTime = 7
Const ColRepType = 8
Const ColRepFile = 9
Const ColKwtDate = 10
Const ColKwtTime = 11
Const ColKwtCode = 12
Const ColKwtNote = 13
Const ColKwtAgain = 14
Const ColKwtAgainTime = 15
Const ColKwtAgainCode = 16
Const ColKwtAgainNote = 17
Const ColKwtAgain2 = 18
Const ColKwtAgain2Time = 19
Const ColKwtAgain2Code = 20
Const ColKwtAgain2Note = 21

Const DateFormat = "d/m;@"
Const TimeFormat = "d/m h:mm;@"

Const ColorGrey = -2236963
Const ColorOK = -4165632
Const ColorBad = -16776961
Const ColorToday = 65535

Dim CntMZ As Integer
Dim CntRep As Integer
Dim CntKwt As Integer

Dim Date1 As Date
Dim Date2 As Date

Public Sub Refresh()
    CntMZ = 0
    CntRep = 0
    CntKwt = 0
    PrepareSheet
    
    Dim Answer As Variant
    Answer = "01." & Format(DateAdd("m", -1, Now), "MM.yyyy")
    Answer = InputBox("Дата начала периода" & vbCrLf & "(с прошлого месяца):", "440-П", Answer)
    If Answer = "" Then Exit Sub
    Date1 = CDate(Answer)
    
    Answer = Format(Now, "dd.MM.yyyy")
    Answer = InputBox("Дата конца периода" & vbCrLf & "(по сегодня):", "440-П", Answer)
    If Answer = "" Then Exit Sub
    Date2 = CDate(Answer)
    
    ActiveCell.Worksheet.Name = "За " & Format(Date1, "dd.MM") & "-" & Format(Date2, "dd.MM")
    
    F440Pin_List
    F440Prep_List
    F440Pkwt_List
    
    FinalSheet
    MsgBox "За период с " & Date1 & " по " & Date2 & vbCrLf & vbCrLf & _
        "Запросов: " & CntMZ & vbCrLf & _
        "Ответов: " & CntRep & vbCrLf & _
        "Квитанций: " & CntKwt, vbInformation, "Статистика 440-П XML"
End Sub

Private Sub PrepareSheet()
    Dim r As Integer
    r = 1
    
    'Cells.Clear
    Cells.Delete Shift:=xlUp
    
    Cells(r, ColID) = "Н/п"
    Cells(r, ColDate) = "Дата"
    Cells(r, ColTime) = "Время"
    Cells(r, ColType) = "Запрос"
    Cells(r, ColMZFile) = "Файл"
    Cells(r, ColRepDate) = "Мы"
    Cells(r, ColRepTime) = "Время"
    Cells(r, ColRepType) = "Ответ"
    Cells(r, ColRepFile) = "Файл"
    Cells(r, ColKwtDate) = "Квит."
    Cells(r, ColKwtTime) = "Время"
    Cells(r, ColKwtCode) = "Код"
    Cells(r, ColKwtNote) = "Примечание"
    Cells(r, ColKwtAgain) = "Повт."
    Cells(r, ColKwtAgainTime) = "Время"
    Cells(r, ColKwtAgainCode) = "Код"
    Cells(r, ColKwtAgainNote) = "Примечание"
    Cells(r, ColKwtAgain2) = "Повт."
    Cells(r, ColKwtAgain2Time) = "Время"
    Cells(r, ColKwtAgain2Code) = "Код"
    Cells(r, ColKwtAgain2Note) = "Примечание"
    
    Columns(ColDate).NumberFormat = DateFormat
    Columns(ColTime).NumberFormat = TimeFormat
    Columns(ColRepDate).NumberFormat = DateFormat
    Columns(ColRepTime).NumberFormat = TimeFormat
    Columns(ColKwtDate).NumberFormat = DateFormat
    Columns(ColKwtTime).NumberFormat = TimeFormat
    Columns(ColKwtAgain).NumberFormat = DateFormat
    Columns(ColKwtAgainTime).NumberFormat = TimeFormat
    Columns(ColKwtAgain2).NumberFormat = DateFormat
    Columns(ColKwtAgain2Time).NumberFormat = TimeFormat
End Sub

Private Sub FinalSheet()
    Cells.AutoFilter
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    ActiveCell.Worksheet.Name = ActiveCell.Worksheet.Name & _
        " на " & Format(Now, "dd.MM HH.mm")
End Sub

Private Sub F440Pin_List()
    Dim d As Date: d = Date1
    Dim r As Integer: r = 2
    
    Dim StrDate As String
    Dim StrPath As String
    Dim StrFile As String
    
    Application.StatusBar = "Посылки..."
    Do
        StrDate = Format(d, DatePath)
        StrPath = F440Pin & StrDate
        StrFile = Dir(StrPath & "*.xml")
        Do While Len(StrFile) > 0
            If Left(StrFile, 3) <> "IZV" And Left(StrFile, 3) <> "KWT" Then
                CntMZ = CntMZ + 1
                Cells(r, ColID) = CntMZ
                Cells(r, ColDate) = d
                Cells(r, ColTime) = FileDateTime(StrPath & StrFile)
                Cells(r, ColType) = Left(StrFile, 3)
                Cells(r, ColMZFile) = StrFile
                r = r + 1
            End If
            StrFile = Dir
        Loop
        
        d = DateAdd("d", 1, d)
        If d > Date2 Then Exit Do
        'DoEvents
    Loop
    
    Columns(ColID).EntireColumn.AutoFit
    Columns(ColDate).EntireColumn.AutoFit
    Columns(ColTime).EntireColumn.AutoFit
    Columns(ColType).EntireColumn.AutoFit
    
    Columns(ColDate).HorizontalAlignment = xlCenter
    Columns(ColTime).HorizontalAlignment = xlCenter
    Columns(ColType).HorizontalAlignment = xlCenter
    
    Application.StatusBar = False
    DoEvents
End Sub

Private Sub F440Prep_List()
    Dim d As Date
    Dim r As Integer
    Dim c As Integer
       
    Dim StrDate As String
    Dim StrPath As String
    Dim StrFile As String
    Dim StrFind As String
    
    Application.StatusBar = "Ответы..."
    
    r = 2
    Do While Len(Cells(r, ColMZFile).Text) > 0
        d = Cells(r, ColDate)
        StrFind = "*" & Replace(Cells(r, ColMZFile).Text, ".xml", "*.*")
        Do
            StrDate = Format(d, DatePath)
            StrPath = F440Prep & StrDate
            StrFile = Dir(StrPath & StrFind)
            Do While Len(StrFile) > 0
                CntRep = CntRep + 1
                r = r + 1
                Rows(r).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                
                For c = ColID To ColMZFile
                    Cells(r, c) = Cells(r - 1, c)
                    Cells(r, c).Font.Color = ColorGrey
                Next
                
                Cells(r, ColRepDate) = d
                Cells(r, ColRepTime) = FileDateTime(StrPath & StrFile)
                If DateDiff("d", d, Now) = 0 Then
                    Cells(r, ColRepDate).Interior.Color = ColorToday
                End If
                Cells(r, ColRepType) = Left(StrFile, 3)
                Cells(r, ColRepFile) = StrFile
                
                Cells(r, ColKwtNote) = "ждем..."
                
                StrFile = Dir
            Loop
            d = DateAdd("d", 1, d)
            If d > Now Then Exit Do
        Loop
        r = r + 1
        'If r > 10 Then Exit Do
        If CntRep Mod 11 = 0 Then
            Application.StatusBar = "Ответы " & CntRep
            DoEvents
        End If
    Loop
    
    Columns(ColRepDate).EntireColumn.AutoFit
    Columns(ColRepTime).EntireColumn.AutoFit
    Columns(ColRepType).EntireColumn.AutoFit
    Columns(ColRepFile).EntireColumn.AutoFit
    
    Columns(ColRepFile).ColumnWidth = Columns(ColRepFile).ColumnWidth * 3 / 4
    
    Columns(ColRepDate).HorizontalAlignment = xlCenter
    Columns(ColRepTime).HorizontalAlignment = xlCenter
    Columns(ColRepType).HorizontalAlignment = xlCenter
   
    Application.StatusBar = False
    DoEvents
End Sub

Private Sub F440Pkwt_List()
    Dim d As Date
    Dim r As Integer
    Dim c As Integer
       
    Dim StrDate As String
    Dim StrPath As String
    Dim StrFile As String
    Dim StrFind As String
    
    Dim XDoc As Object, root As Object, node As Object
    
    r = 2
    Do While Len(Cells(r, ColID).Text) > 0
        If Len(Cells(r, ColRepFile).Text) > 0 Then
            c = ColKwtDate
            d = Cells(r, ColRepDate)
            d = DateAdd("d", 1, d) 'квитанции из ФНС не могут придти в тот же день - защита от квитования с предыдущей квитанцией
            StrFind = "KWT*" & Cells(r, ColRepFile).Text
            Do
                StrDate = Format(d, DatePath)
                StrPath = F440Pin & StrDate
                StrFile = Dir(StrPath & StrFind)
                Do While Len(StrFile) > 0
                    CntKwt = CntKwt + 1
                    Cells(r, c) = d
                    Cells(r, c + 1) = FileDateTime(StrPath & StrFile)
                    If DateDiff("d", d, Now) = 0 Then
                        Cells(r, c).Interior.Color = ColorToday
                    End If
                    c = c + 2
                    
                    Set XDoc = CreateObject("Microsoft.XMLDOM")
                    XDoc.async = False
                    XDoc.validateonparse = False
                    XDoc.Load (StrPath & StrFile)
                    
                    Set node = XDoc.SelectSingleNode("/Файл/КВТНОПРИНТ/Результат")
                    'Cells(r, c) = node.Attributes("КодРезПроверки").Text
                    If node.Attributes(0).Text = "01" Then
                        Cells(r, ColRepFile).Font.Color = ColorOK
                        Cells(r, c) = "'01" 'node.Attributes(0).Text
                        c = c + 1
                        Cells(r, c) = "OK"
                        Cells(r, c).Font.Color = ColorOK
                        c = c + 1
                    Else
                        Cells(r, ColRepFile).Font.Color = ColorBad
                        Cells(r, c) = "'" & node.Attributes(0).Text
                        c = c + 1
                        'Cells(r, c) = node.Attributes("Пояснение").Text
                        Cells(r, c) = node.Attributes(1).Text
                        Cells(r, c).Font.Color = ColorBad
                        c = c + 1
                    End If
                    
                    Set node = Nothing
                    Set root = Nothing
                    StrFile = Dir
                Loop
                d = DateAdd("d", 1, d)
                If d > Now Then Exit Do
            Loop
        End If
        r = r + 1
        'If r > 10 Then Exit Do
        If CntKwt Mod 11 = 0 Then
            Application.StatusBar = "Квитанции " & CntKwt
            DoEvents
        End If
    Loop
    Set XDoc = Nothing
    
    Columns(ColKwtDate).EntireColumn.AutoFit
    Columns(ColKwtTime).EntireColumn.AutoFit
    Columns(ColKwtCode).EntireColumn.AutoFit
    Columns(ColKwtAgain).EntireColumn.AutoFit
    Columns(ColKwtAgainTime).EntireColumn.AutoFit
    Columns(ColKwtAgainCode).EntireColumn.AutoFit
    
    Columns(ColKwtNote).ColumnWidth = Columns(ColType).ColumnWidth
    
    Columns(ColKwtDate).HorizontalAlignment = xlCenter
    Columns(ColKwtTime).HorizontalAlignment = xlCenter
    Columns(ColKwtCode).HorizontalAlignment = xlCenter
    Columns(ColKwtAgain).HorizontalAlignment = xlCenter
    Columns(ColKwtAgainTime).HorizontalAlignment = xlCenter
    Columns(ColKwtAgainCode).HorizontalAlignment = xlCenter
    
    Application.StatusBar = False
    DoEvents
End Sub
