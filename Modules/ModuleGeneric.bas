Attribute VB_Name = "ModuleGeneric"
Option Explicit
Option Base 1

'Standard μεταβλητές
Global Const strAppTitle = "Wheels and Tires v3"
Global strApplicationEXEName As String
Global arrCompanyData(10) As String
Global arrData(13) As String
Global arrMenu() As Integer
Global strBankAccountNo As String
Global strInvoiceExtraRemarksA As String
Global strInvoiceExtraRemarksB As String
Global blnUseNewInvoiceForm As Boolean

'Databases
Global strDatabaseName As String
Global strDatabaseType As String
Global strDatabasePort As String
Global strDatabaseServer As String
Global CommonDB As Database
Global PrintersDB As Database
Global UsersDB As Database
Global dBaseTables As TableDefs
Global TempQuery As QueryDef

'Εκτυπωτές
Global strPrinterName As String
Global strPrinterData(5) As String

'Μεταβλητές
Global strMessages(100) As String
Global strCurrentUser As String
Global strFullPathName As String
Global strPathName As String
Global strReportsPathName As String
Global strCompanyName As String
Global strImageDirectory As String
Global strUnicodeFile As String
Global strAsciiFile As String
Global blnAppIsRunning As Boolean
Global blnPreviewInvoices As Boolean
Global blnPreviewReports As Boolean

'Indexes
Public Type typTableData
    strCode As String
    strOneField As String
    strTwoField As String
    strThreeField As String
    strFourField As String
    strFiveField As String
    strSixField As String
    strSevenField As String
    strEightField As String
    strNineField As String
    strTenField As String
    strElevenField As String
    strTwelveField As String
    strThirteenField As String
    strFourteenField As String
    strFifteenField As String
    strSixteenField As String
    strSeventeenField As String
End Type

Function ToggleFieldVisibility(visibility As Boolean, ParamArray tmpFields())

    Dim bytLoop As Byte
    
    For bytLoop = 0 To UBound(tmpFields)
        tmpFields(bytLoop).ForeColor = IIf(visibility = True, vbBlack, vbWhite)
    Next bytLoop

End Function

Function KillProcess(appName)

    Dim process As Object

    For Each process In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process")
        If process.Caption = appName Then
            process.Terminate (0)
        End If
    Next

End Function


Function CheckForInteger(myText)

    If Not IsNumeric(myText) Then CheckForInteger = False Else CheckForInteger = True

End Function

Public Function FillCellWithSomething(myGrid As iGrid, myFiller, myCurrentRow, myColumns)

    Dim lngRow As Long
    Dim lngCol As Long
    Dim myCols() As String
    
    myCols = Split(myColumns, ",")
    
    myGrid.Redraw = False
    
    If myCurrentRow <> 0 Then
        For lngCol = 0 To UBound(myCols)
            myGrid.CellValue(myCurrentRow, Val(myCols(lngCol))) = myFiller
        Next lngCol
    Else
        For lngRow = 1 To myGrid.RowCount
            For lngCol = 1 To UBound(myCols)
                myGrid.CellValue(lngRow, Val(myCols(lngCol))) = myFiller
            Next lngCol
        Next lngRow
    End If
    
    myGrid.Redraw = True

End Function

Function FindUTF8Char(strChar)

    Dim c As Long
    Dim n As Integer
    Dim utftext As String
  
    utftext = ""
    
    For n = 1 To Len(strChar)
        c = AscW(Mid(strChar, n, 1))
        If c < 128 Then
            utftext = utftext + Mid(strChar, n, 1)
        ElseIf ((c > 127) And (c < 2048)) Then
            utftext = utftext + Chr(((c \ 64) Or 192))
            utftext = utftext + Chr(((c And 63) Or 128))
        Else
            utftext = utftext + Chr(((c \ 144) Or 234))
            utftext = utftext + Chr((((c \ 64) And 63) Or 128))
            utftext = utftext + Chr(((c And 63) Or 128))
        End If
    Next n

  FindUTF8Char = utftext

End Function

Function AddGridLines(myGrid As iGrid, myRefersTo, myLines As Integer)

    myGrid.AddRow , , , , , , , IIf(myRefersTo = "1", 50, myLines)

End Function

Function CheckForAcceptableKey(myKeyCode)

    CheckForAcceptableKey = IIf((myKeyCode >= 48 And myKeyCode <= 57) Or myKeyCode = 46 Or myKeyCode = 44 Or myKeyCode = 45 Or myKeyCode = 8 Or myKeyCode = 13, True, False)

End Function

Function CheckForMaxLength(myText, myMaxLength, Optional myFormat)

    If myFormat = "Float" Then
        myText = format(myText, "#,##0.00")
    End If
    
    CheckForMaxLength = IIf(Len(myText) <= myMaxLength, myText, "")

End Function


Function CheckToEnableButton(myGrid As iGrid, myLine, myColumn, Optional myLoadedForm As String)

    On Error GoTo ErrTrap
    
    If myLine > 0 Then CheckToEnableButton = IIf(myGrid.CellText(myLine, myColumn) <> "", True, False)
    
    If CheckForLoadedForm(myLoadedForm) Then CheckToEnableButton = False
    
    Exit Function
    
ErrTrap:
    DisplayErrorMessage True, Err.Description

End Function

Function OldClearGridCell(myGrid As iGrid, myRow As Long, Optional myCol As Long)

    Dim lngLoop As Long
    
    If myCol = 0 Then
        For lngLoop = 1 To myGrid.colCount
            myGrid.CellValue(myRow, lngLoop) = ""
        Next lngLoop
        Exit Function
    End If
    
End Function

Function ClearVariables(ParamArray myFields() As Variant)

    Dim intLoop As Integer
    
    For intLoop = 0 To UBound(myFields)
        If VarType(myFields) = 8204 Then
            myFields(intLoop) = 0
        End If
    Next intLoop

End Function

Function ColorizeControls(thisForm As Form, Optional fullScreen As Boolean, Optional customColours As Boolean)

    Dim ctl As Control
    Dim objFont As StdFont
    
    If Not customColours Then
        thisForm.BackColor = IIf(fullScreen, GetSetting(strAppTitle, "Appearance", "Background Full Screen Forms"), GetSetting(strAppTitle, "Appearance", "Forms Centered Background"))
    End If
    
    For Each ctl In thisForm.Controls
        'Κριτήρια
        If ctl.Name = "frmCriteria" Then
            ctl.BackColor = GetSetting(strAppTitle, "Appearance", "Background Criteria")
        End If
        'Container
        If ctl.Name = "frmContainer" Then
            ctl.BackColor = IIf(fullScreen, GetSetting(strAppTitle, "Appearance", "Forms FullScreen Background"), GetSetting(strAppTitle, "Appearance", "Background Containers"))
        End If
        'Φόντο
        If ctl.Name = "shpBackground" Then
            ctl.BackColor = IIf(fullScreen, GetSetting(strAppTitle, "Appearance", "Forms FullScreen Background"), GetSetting(strAppTitle, "Appearance", "Frames Background"))
        End If
        'Πλαίσιο κουμπιών
        If ctl.Name = "frmButtonFrame" Or ctl.Name = "frmFrameForGridButtons" Or ctl.Name = "frmTotals" Or ctl.Name = "frmDetails" Then
            ctl.BackColor = thisForm.BackColor
        End If
        'Πλέγμα
        If TypeOf ctl Is iGrid And Not customColours Then
            ctl.BackColor = IIf(fullScreen, GetSetting(appName:=strAppTitle, Section:="Appearance", Key:="Grid FullScreen BackColor"), GetSetting(appName:=strAppTitle, Section:="Appearance", Key:="Grid BackColor"))
            ctl.GridLines = IIf(fullScreen, GetSetting(appName:=strAppTitle, Section:="Appearance", Key:="Grid FullScreen GridLines"), GetSetting(appName:=strAppTitle, Section:="Appearance", Key:="Grid GridLines"))
            ctl.ForeColor = IIf(fullScreen, GetSetting(appName:=strAppTitle, Section:="Appearance", Key:="Grid FullScreen ForeColor"), GetSetting(appName:=strAppTitle, Section:="Appearance", Key:="Grid ForeColor"))
            ctl.HighlightForeColor = IIf(fullScreen, GetSetting(appName:=strAppTitle, Section:="Appearance", Key:="Grid FullScreen Highlight ForeColor"), GetSetting(appName:=strAppTitle, Section:="Appearance", Key:="Grid Highlight ForeColor"))
            ctl.HighlightBackColor = IIf(fullScreen, GetSetting(appName:=strAppTitle, Section:="Appearance", Key:="Grid FullScreen Highlight BackColor"), GetSetting(appName:=strAppTitle, Section:="Appearance", Key:="Grid Highlight BackColor"))
        End If
        'Ετικέτες
        If TypeOf ctl Is Label Then
            Select Case ctl.Name
                'Ετικέτα σε φόρμα όχι πλήρους οθόνης
                Case "lblLabel"
                    ctl.ForeColor = GetSetting(strAppTitle, "Appearance", "Labels Normal Foreground")
                    ctl.BackStyle = 0
                'Ετικέτα σε πλαίσιο κριτηρίων
                Case "lblCriteriaLabel"
                    ctl.ForeColor = GetSetting(strAppTitle, "Appearance", "Labels Criteria Foreground")
                    ctl.BackStyle = 0
                Case "lblSimple"
                    ctl.ForeColor = vbWhite
                    ctl.BackStyle = 0
                    Set objFont = New StdFont
                    objFont.Name = GetSetting(strAppTitle, "Appearance", "Labels Title Font")
                    objFont.Size = 10
                    objFont.Bold = False
                    Set ctl.Font = objFont
            End Select
        End If
        'Ετικέτες επικεφαλίδων φόρμας
        If TypeOf ctl Is Label Then
            Select Case ctl.Name
                'Ετικέτες τίτλου
                Case "lblTitle"
                    If Not customColours Then
                        ctl.ForeColor = GetSetting(strAppTitle, "Appearance", "Labels Title Foreground")
                    End If
                    Set objFont = New StdFont
                    objFont.Name = GetSetting(strAppTitle, "Appearance", "Labels Title Font")
                    objFont.Size = 30
                    objFont.Bold = True
                    objFont.Charset = 161
                    Set ctl.Font = objFont
                    Set objFont = Nothing
                Case "lblCriteria"
                    ctl.ForeColor = GetSetting(strAppTitle, "Appearance", "Labels Totals Criteria")
            End Select
        End If
        
        'Checkboxes
        If TypeOf ctl Is CheckBox And Not customColours Then
            'Checkbox σε φόρμα
            If Left(ctl.Name, 11) <> "chkCriteria" Then
                ctl.ForeColor = GetSetting(strAppTitle, "Appearance", "Checkbox Normal Foreground")
                ctl.BackColor = GetSetting(strAppTitle, "Appearance", "Checkbox Normal Background")
            End If
            'Checkbox σε πλαίσιο κριτηρίων
            If Left(ctl.Name, 11) = "chkCriteria" Then
                ctl.ForeColor = GetSetting(strAppTitle, "Appearance", "Labels Criteria Foreground")
                ctl.BackColor = GetSetting(strAppTitle, "Appearance", "Background Criteria")
            End If
        End If
        
        'Radios
        If TypeOf ctl Is OptionButton And Not customColours Then
            'Radios σε φόρμα
            If Left(ctl.Name, 11) <> "optCriteria" Then
                ctl.ForeColor = GetSetting(strAppTitle, "Appearance", "OptionButton Normal Foreground")
                ctl.BackColor = GetSetting(strAppTitle, "Appearance", "OptionButton Normal Background")
            End If
            'Radios σε πλαίσιο κριτηρίων
            If Left(ctl.Name, 11) = "optCriteria" Then
                ctl.ForeColor = GetSetting(strAppTitle, "Appearance", "Labels Criteria Foreground")
                ctl.BackColor = GetSetting(strAppTitle, "Appearance", "Labels Criteria Background")
            End If
        End If
        
        'Frames
        If TypeOf ctl Is Frame And Not customColours Then
            If ctl.Tag = "SameColorAsBackground" Then
                ctl.ForeColor = GetSetting(strAppTitle, "Appearance", "Frames Foreground")
                ctl.BackColor = GetSetting(strAppTitle, "Appearance", "Frames Background")
            End If
        End If
        
    Next
    
End Function

Function ConvertCharacterToUpperCase(strCharacter)

    Select Case Asc(strCharacter)
        Case 220
            ConvertCharacterToUpperCase = "Α"
        Case 221
            ConvertCharacterToUpperCase = "Ε"
        Case 222
            ConvertCharacterToUpperCase = "Η"
        Case 223
            ConvertCharacterToUpperCase = "Ι"
        Case 252
            ConvertCharacterToUpperCase = "Ο"
        Case 253
            ConvertCharacterToUpperCase = "Υ"
        Case 254
            ConvertCharacterToUpperCase = "Ω"
        Case 162
            ConvertCharacterToUpperCase = "Α"
        Case 184
            ConvertCharacterToUpperCase = "Ε"
        Case 185
            ConvertCharacterToUpperCase = "Η"
        Case 186
            ConvertCharacterToUpperCase = "Ι"
        Case 188
            ConvertCharacterToUpperCase = "Ο"
        Case 190
            ConvertCharacterToUpperCase = "Υ"
        Case 191
            ConvertCharacterToUpperCase = "Ω"
        Case 218
            ConvertCharacterToUpperCase = "Ι"
        Case 219
            ConvertCharacterToUpperCase = "Υ"
        Case 250
            ConvertCharacterToUpperCase = "Ι"
        Case 251
            ConvertCharacterToUpperCase = "Υ"
        Case 242
            ConvertCharacterToUpperCase = "Σ"
        Case Else
            ConvertCharacterToUpperCase = UCase(strCharacter)
    End Select
        
End Function

Function CountSelected(myGrid As iGrid)

    Dim lngRow As Long
    Dim intSelected As Integer
    
    For lngRow = 1 To myGrid.RowCount
        If myGrid.CellIcon(lngRow, "Selected") > 0 Then
            intSelected = intSelected + 1
        End If
    Next lngRow
    
    CountSelected = IIf(intSelected > 0, "Επιλεγμένες " & intSelected & " εγγραφές", "")

End Function

Function CreatePDF(myPaperSize, myOrientation, myTopMargin, myLeftMargin, myWindowTitle, myFontName, myFontSize, myInvoiceOrReport)

    If myInvoiceOrReport = "PrinterPrintsInvoicesID" Then
        
        strPrinterData(1) = myPaperSize
        strPrinterData(2) = myOrientation
        strPrinterData(3) = myTopMargin
        strPrinterData(4) = myLeftMargin
        
        With rptInvoiceA
            .Restart
            .Caption = ""
            .PageSettings.PaperSize = myPaperSize
            .PageSettings.Orientation = myOrientation
            .PageSettings.LeftMargin = myLeftMargin * 100
            .PageSettings.TopMargin = myTopMargin * 100
        End With
    
    End If

    If myInvoiceOrReport = "PrinterPrintsReportsID" Then
        
        strPrinterData(1) = myPaperSize
        strPrinterData(2) = myOrientation
        strPrinterData(3) = myTopMargin
        strPrinterData(4) = myLeftMargin
        
        With rptReport
            .Restart
            .Caption = "ΠΡΟΕΠΙΣΚΟΠΗΣΗ: " & CustomUpperCase(myWindowTitle)
            .PageSettings.PaperSize = myPaperSize
            .PageSettings.Orientation = myOrientation
            .PageSettings.LeftMargin = myLeftMargin * 100
            .PageSettings.TopMargin = myTopMargin * 100
            With .oneLongField
                .Width = 10100
                .Font.Name = myFontName
                .Font.Size = myFontSize
            End With
        End With
    
    End If

End Function

Function FormatAsSelection(myFormat)

    Select Case myFormat
        Case "F"
            FormatAsSelection = "#,#0.00"
        Case "I"
            FormatAsSelection = "#,#0"
    End Select

End Function

Function NewCheckForMatch(DBToUse, myGivenFields, myTable, myJoins, myCriteria, myGroupByColumns, myOrderColumns) As Recordset

    On Error GoTo ErrTrap
    
    Dim strSQL As String
    
    Dim rstTempRecordset As Recordset
    
    If DBToUse = "PrintersDB" Then
        Set PrintersDB = DBEngine.OpenDataBase(App.Path + "\" + "Data" + "\" + "Printers.mdb", False, False)
        Set TempQuery = PrintersDB.CreateQueryDef("")
    Else
        Set TempQuery = CommonDB.CreateQueryDef("")
    End If
    
    If Left(myCriteria, 1) = "*" Then
        myCriteria = "1 = 1"
    End If
    
    strSQL = "SELECT " & IIf(myGivenFields = "", "*", myGivenFields) & " FROM " & myTable & " " & myJoins & IIf(myCriteria <> "", " WHERE " & myCriteria, "") & IIf(myGroupByColumns <> "", myGroupByColumns, "") & IIf(myOrderColumns <> "", " ORDER BY " & myOrderColumns, "")
    
    TempQuery.SQL = strSQL
    
    Set rstTempRecordset = TempQuery.OpenRecordset()
    Set NewCheckForMatch = rstTempRecordset
    
    Exit Function
    
ErrTrap:
    DisplayErrorMessage True, Err.Description

End Function

Function PrintPDF(myPrinterName)

    With rptReport
        .Printer.DeviceName = myPrinterName
        .PrintReport False
        .Run True
    End With

End Function

Function PrintInvoiceToLaser(myInvoiceTrnID, myPrinterName)

    Dim intLoop As Integer
    Dim pdf As New ARExportPDF
    
    For intLoop = 1 To 2
        rptInvoiceA.Restart
        rptInvoiceA.Tag = myInvoiceTrnID
        rptInvoiceA.PageSettings.Orientation = ddOLandscape
        rptInvoiceA.PageSettings.PaperSize = 11
        rptInvoiceA.lblIsOriginalOrCopy.Caption = IIf(intLoop = 1, "ΠΡΩΤΟΤΥΠΟ", "ΑΝΤΙΓΡΑΦΟ")
        rptInvoiceA.PageSettings.LeftMargin = 500
        rptInvoiceA.PageSettings.RightMargin = 50
        rptInvoiceA.Run False
        pdf.AcrobatVersion = 2
        pdf.SemiDelimitedNeverEmbedFonts = ""
        pdf.fileName = strReportsPathName & "ΤΙΜΟΛΟΓΙΟ " & rptInvoiceA.lblIsOriginalOrCopy.Caption & ".pdf"
        pdf.Export rptInvoiceA.Pages
        'If blnPreviewInvoices Then
        '    rptInvoiceA.Zoom = -2
        '    rptInvoiceA.Printer.ColorMode = vbPRCMMonochrome
        '    rptInvoiceA.WindowState = vbMaximized
        '    rptInvoiceA.Show 1
        'Else
        '    rptInvoiceA.Printer.DeviceName = myPrinterName
        '    rptInvoiceA.PrintReport False
        'End If
    Next intLoop

End Function

Function SumSelectedGridRows(myGrid As iGrid, myLastColumnIsSpecial, ParamArray myColumns() As Variant)

    Dim lngRow As Long
    Dim intLoop As Integer
    Dim blnSelected As Boolean
    Dim strDummy As String
    ReDim curGridColumnTotals(UBound(myColumns) + 1)
    
    For lngRow = 1 To myGrid.RowCount
        If myGrid.CellIcon(lngRow, "Selected") > 0 Then
            blnSelected = True
            For intLoop = 1 To UBound(myColumns) + IIf(myLastColumnIsSpecial, 0, 1)
                curGridColumnTotals(intLoop) = curGridColumnTotals(intLoop) + myGrid.CellValue(lngRow, myColumns(intLoop - 1))
            Next intLoop
            If intLoop - 1 = UBound(myColumns) And myLastColumnIsSpecial Then
                curGridColumnTotals(intLoop) = curGridColumnTotals(intLoop) + myGrid.CellValue(lngRow, myColumns(intLoop - 3)) - myGrid.CellValue(lngRow, myColumns(intLoop - 2))
            End If
        End If
    Next lngRow
    
    If blnSelected Then
        For intLoop = 1 To UBound(myColumns) + 1
            strDummy = strDummy & myGrid.ColHeaderText(myColumns(intLoop - 1)) & " " & format(curGridColumnTotals(intLoop), "#,##0.00") & " "
        Next intLoop
        SumSelectedGridRows = Left(strDummy, Len(strDummy) - 1)
    End If

End Function

Function DisplayMessage(myMessageNumber, myIcon, myChoices, mySuffix, ParamArray myFieldsToCheck() As Variant)

    'Λάθος πεδίο
    If UBound(myFieldsToCheck()) = -1 Then
        If MyMsgBox(myIcon, strAppTitle, strMessages(myMessageNumber) & mySuffix, myChoices) Then
        End If
        DisplayMessage = True
        Exit Function
    End If

    'Υποχρεωτικό πεδίο
    If Len(myFieldsToCheck(0)) = 0 And UBound(myFieldsToCheck()) = 0 Then
        If MyMsgBox(myIcon, strAppTitle, strMessages(myMessageNumber) & mySuffix, myChoices) Then
        End If
        DisplayMessage = True
        Exit Function
    End If
    
    'Ημερομηνίες Έως >= Από
    If UBound(myFieldsToCheck()) = 1 Then
        'Αν έχω δώσει "Από"
        If Len(myFieldsToCheck(0)) > 0 Then
            'Αν έχω δώσει και τις δύο ημερομηνίες
            If Len(myFieldsToCheck(0)) = 10 And Len(myFieldsToCheck(1)) = 10 Then
                'Ελέγχω για σωστό διάστημα
                If CDate(myFieldsToCheck(0)) > CDate(myFieldsToCheck(1)) Then
                    If MyMsgBox(myIcon, strAppTitle, strMessages(myMessageNumber) & mySuffix, myChoices) Then
                    End If
                    DisplayMessage = True
                    Exit Function
                End If
            End If
        End If
    End If
    
End Function

Function HighlightNextRow(grdGrid As iGrid, lngRow, lngColumn, blnRowMode)

    With grdGrid
        If .RowCount > 0 Then
            If lngRow = 1 Then
                .EnsureVisibleRow 1
                .SetCurCell 1, lngColumn
                .RowMode = blnRowMode
                .SetFocus
                Exit Function
            End If
            If lngRow <= grdGrid.RowCount Then
                .EnsureVisibleRow lngRow
                .SetCurCell lngRow, lngColumn
                .RowMode = blnRowMode
                .SetFocus
            Else
                .EnsureVisibleRow .RowCount
                .SetCurCell .RowCount, lngColumn
                .RowMode = blnRowMode
                .SetFocus
            End If
        End If
    End With

End Function

Function DisableTabStop(ParamArray tmpFields())
    
    Dim bytLoop As Byte
    
    For bytLoop = 0 To UBound(tmpFields)
        tmpFields(bytLoop).TabStop = False
    Next bytLoop

End Function

Function IsCorrectPassword(strUserName, strPassword As String)

    Dim rstUsers As Recordset
    Dim strUserInput As String
    
    Set UsersDB = DBEngine.OpenDataBase(strPathName + "Users.mdb", False, False)
    Set TempQuery = UsersDB.CreateQueryDef("")
    
    TempQuery.SQL = "SELECT * FROM Users WHERE Username = '" & strUserName & "' AND PasswordHash = '" & HashPassword(strUserName, strPassword) & "'"
    
    Set rstUsers = TempQuery.OpenRecordset()
    
    If Not rstUsers.EOF Then
        IsCorrectPassword = True
    Else
        IsCorrectPassword = False
    End If
    
    UsersDB.Close
    
End Function

Function OnlyAcceptSpecificValues(myInput, ParamArray myValues() As Variant)

    Dim intLoop As Integer
    
    For intLoop = 0 To UBound(myValues)
        If myInput = myValues(intLoop) Then OnlyAcceptSpecificValues = True: Exit Function
    Next intLoop

End Function

Function PositionCenteredScreenControls(thisForm As Form, formFullScreen As Boolean, Optional grdGrid As iGrid, Optional customColours As Boolean)

    'Φόρμα
    thisForm.Width = thisForm.shpRightEdge.Left + thisForm.shpRightEdge.Width
    thisForm.Height = thisForm.shpBottomEdge.Top + thisForm.shpBottomEdge.Height - 90
    thisForm.Left = CommonMain.Width / 2 - thisForm.Width / 2 - 100
    thisForm.Top = CommonMain.Height / 2 - thisForm.Height / 2
    
    'Κουμπιά
    With thisForm.frmButtonFrame
        .Left = (thisForm.Width / 2) - (thisForm.frmButtonFrame.Width / 2)
    End With
    
    'Τετράγωνο πλαίσιο
    With thisForm.shpBackground
        .Top = 900
        .Left = 225
        .Width = thisForm.Width - 470
        .Height = thisForm.frmButtonFrame.Top - 270 - .Top
    End With
   
End Function

Function CustomizeGrid(ParamArray myGrid() As Variant)

    Dim intLoop As Integer
    
    For intLoop = 0 To UBound(myGrid)
        With myGrid(intLoop)
            .GridLineColor = GetSetting(appName:=strAppTitle, Section:="Appearance", Key:="Grid Header BackColor")
            .GridLines = igGridLinesBoth
            .GridLinesExtend = igGridLinesExtendDown
            .RowMode = False
            .TabStop = False
        End With
    Next intLoop

End Function

Function PositionFullScreenControls(thisForm As Form, formFullScreen As Boolean, Optional grdGrid As iGrid, Optional customColours As Boolean)

    Dim ctl As Control
    
    'Φόρμα
    With thisForm
        .Top = 350
        .Height = CommonMain.Height - (.Top * 1.2)
        .Width = CommonMain.Width
        .Left = -100
    End With
    
    'Container
    With thisForm.frmContainer
        .Height = thisForm.Height - 510
        .Top = (thisForm.Height / 2) - (.Height / 2)
        .Left = (thisForm.Width / 2) - (.Width / 2)
    End With
    
    'Κουμπιά
    With thisForm.frmButtonFrame
        .Top = thisForm.frmContainer.Height - 840
        .Left = (thisForm.frmContainer.Width / 2) - (.Width / 2)
    End With
    
    'Τετράγωνο πλαίσιο
    With thisForm.shpBackground
        .Top = 975
        .Left = 0
        .Width = thisForm.Width
        .Height = thisForm.frmButtonFrame.Top - 200 - .Top
    End With
    
    'Πλέγμα
    With grdGrid
        .Height = thisForm.shpBackground.Height + 180 - .Top + (thisForm.Top * 2)
        .ForeColor = vbWhite
        .HighlightForeColor = vbBlack
        .HighlightBackColor = &HC0FFC0
    End With
    
    For Each ctl In thisForm.Controls
        'Κουμπιά που αφορούν το πλέγμα
        If ctl.Name = "frmFrameForGridButtons" Then
            With thisForm.frmFrameForGridButtons
                .Top = thisForm.shpBackground.Height + 550
                .Left = (thisForm.frmContainer.Width / 2) - (.Width / 2)
            End With
            grdGrid.Height = thisForm.Height - 2050 - thisForm.frmFrameForGridButtons.Height
        End If
        'Σύνολα αγορών - πωλήσεων
        If ctl.Name = "frmTotals" Then
            With thisForm.frmTotals
                .Top = thisForm.shpBackground.Height - 190
                .Left = (thisForm.frmContainer.Width / 2) - (.Width / 2)
            End With
            With thisForm.frmDetails
                .Top = thisForm.frmTotals.Top - .Height - 90
                .Left = (thisForm.frmContainer.Width / 2) - (.Width / 2)
            End With
            grdGrid.Height = thisForm.Height - 6190 - thisForm.frmDetails.Height
        End If
    Next ctl
    
    'Κριτήρια
    Dim intIndex As Integer
    intIndex = 0
    For Each ctl In thisForm.Controls
        If Left(ctl.Name, 11) = "frmCriteria" Then
            With thisForm.frmCriteria(intIndex)
                .Visible = True
                .ZOrder 0
                .Top = ((grdGrid.Height) / 2) - (.Height / 2) + grdGrid.Top
                .Left = (grdGrid.Width / 2) - (.Width / 2)
                intIndex = intIndex + 1
            End With
        End If
    Next ctl
   
End Function

Function ScanGridForSelectedRecords(myGrid As iGrid)

    Dim lngRow As Long
    
    For lngRow = 1 To myGrid.RowCount
        If myGrid.CellIcon(lngRow, "Selected") > 0 Then
            ScanGridForSelectedRecords = True
            Exit Function
        End If
    Next lngRow
    
    ScanGridForSelectedRecords = False

End Function

Public Function ToHexDump(sText As String) As String
    
    Dim lIdx As Long

    For lIdx = 1 To Len(sText)
        ToHexDump = ToHexDump & Right$("0" & Hex(Asc(Mid(sText, lIdx, 1))), 2)
    Next
    
End Function

Private Function pvCryptXor(ByVal lI As Long, ByVal lJ As Long) As Long
    
    If lI = lJ Then
        pvCryptXor = lJ
    Else
        pvCryptXor = lI Xor lJ
    End If
    
End Function

Public Function CryptRC4(sText, sKey) As String
    
    Dim baS(0 To 255) As Byte
    Dim baK(0 To 255) As Byte
    Dim bytSwap As Byte
    Dim lI As Long
    Dim lJ As Long
    Dim lIdx As Long

    For lIdx = 0 To 255
        baS(lIdx) = lIdx
        baK(lIdx) = Asc(Mid$(sKey, 1 + (lIdx Mod Len(sKey)), 1))
    Next
    
    For lI = 0 To 255
        lJ = (lJ + baS(lI) + baK(lI)) Mod 256
        bytSwap = baS(lI)
        baS(lI) = baS(lJ)
        baS(lJ) = bytSwap
    Next
    
    lI = 0
    lJ = 0
    
    For lIdx = 1 To Len(sText)
        lI = (lI + 1) Mod 256
        lJ = (lJ + baS(lI)) Mod 256
        bytSwap = baS(lI)
        baS(lI) = baS(lJ)
        baS(lJ) = bytSwap
        CryptRC4 = CryptRC4 & Chr$((pvCryptXor(baS((CLng(baS(lI)) + baS(lJ)) Mod 256), Asc(Mid$(sText, lIdx, 1)))))
    Next
    
End Function

Function UpdateLogFile(errorDescription)

    Open strReportsPathName & "Errors.txt" For Append As #2
        Print #2, format(Date, "dd/mm/yyyy") & " " & format(Time, "hh:mm") & " " & errorDescription; ""
    Close #2
    
End Function

Function ChangeEditButtonStatus(grdGrid, lngRow, lngCol)

    ChangeEditButtonStatus = False
    
    If grdGrid.RowCount = 0 Or lngRow = 0 Then Exit Function
    
    If grdGrid.CellValue(lngRow, lngCol) <> "" Then ChangeEditButtonStatus = True

End Function

Function CalculateColumnTotal(myGrid As iGrid, myColumn)

    Dim lngRow As Long
    Dim curTotal As Currency
    
    For lngRow = 1 To myGrid.RowCount
        If myGrid.CellValue(lngRow, myColumn) <> "" Then
            curTotal = curTotal + myGrid.CellValue(lngRow, myColumn)
        End If
    Next lngRow
    
    CalculateColumnTotal = curTotal
    
End Function

Function ConvertToAsciiFile(strUnicodeFile, strAsciiFile)
    
    On Error GoTo ErrTrap
    
    Dim intLoop As Integer
    Dim strLine As String
    Dim strAsciiLine As String
    Dim strEscCharacter As String
    Dim strInputFile As String
    Dim strOutputFile As String
    
    strInputFile = strUnicodeFile
    strOutputFile = strAsciiFile
    
    Open strInputFile For Input As #1
    Open strOutputFile For Output As #2
    
    Do While Not EOF(1)
        Line Input #1, strLine
        For intLoop = 1 To Len(strLine)
            strAsciiLine = strAsciiLine & FindDOSChar(Mid(strLine, intLoop, 1))
            DoEvents
        Next intLoop
        Print #2, strAsciiLine
        strAsciiLine = ""
    Loop
    
    Close #1
    Close #2
    
    ConvertToAsciiFile = strOutputFile
    
    Exit Function
    
ErrTrap:
    Close #1
    Close #2
    ConvertToAsciiFile = "Error"
    DisplayErrorMessage True, Err.Description

End Function

Function FindDOSChar(strChar)

    Select Case Asc(strChar)
    
    Case 193
        FindDOSChar = Chr(128): Rem Α
    Case 194
        FindDOSChar = Chr(129): Rem Β
    Case 195
        FindDOSChar = Chr(130): Rem Γ
    Case 196
        FindDOSChar = Chr(131): Rem Δ
    Case 197
        FindDOSChar = Chr(132): Rem Ε
    Case 198
        FindDOSChar = Chr(133): Rem Ζ
    Case 199
        FindDOSChar = Chr(134): Rem Η
    Case 200
        FindDOSChar = Chr(135): Rem Θ
    Case 201
        FindDOSChar = Chr(136): Rem Ι
    Case 202
        FindDOSChar = Chr(137): Rem Κ
    Case 203
        FindDOSChar = Chr(138): Rem Λ
    Case 204
        FindDOSChar = Chr(139): Rem Μ
    Case 205
        FindDOSChar = Chr(140): Rem Ν
    Case 206
        FindDOSChar = Chr(141): Rem Ξ
    Case 207
        FindDOSChar = Chr(142): Rem Ο
    Case 208
        FindDOSChar = Chr(143): Rem Π
    Case 209
        FindDOSChar = Chr(144): Rem Ρ
    Case 211
        FindDOSChar = Chr(145): Rem Σ
    Case 212
        FindDOSChar = Chr(146): Rem Τ
    Case 213
        FindDOSChar = Chr(147): Rem Υ
    Case 214
        FindDOSChar = Chr(148): Rem Φ
    Case 215
        FindDOSChar = Chr(149): Rem Χ
    Case 216
        FindDOSChar = Chr(150): Rem Ψ
    Case 217
        FindDOSChar = Chr(151): Rem Ω
        
    Case 225
        FindDOSChar = Chr(152): Rem α
    Case 226
        FindDOSChar = Chr(153): Rem β
    Case 227
        FindDOSChar = Chr(154): Rem γ
    Case 228
        FindDOSChar = Chr(155): Rem δ
    Case 229
        FindDOSChar = Chr(156): Rem ε
    Case 230
        FindDOSChar = Chr(157): Rem ζ
    Case 231
        FindDOSChar = Chr(158): Rem η
    Case 232
        FindDOSChar = Chr(159): Rem θ
    Case 233
        FindDOSChar = Chr(160): Rem ι
    Case 234
        FindDOSChar = Chr(161): Rem κ
    Case 235
        FindDOSChar = Chr(162): Rem λ
    Case 236
        FindDOSChar = Chr(163): Rem μ
    Case 237
        FindDOSChar = Chr(164): Rem ν
    Case 238
        FindDOSChar = Chr(165): Rem ξ
    Case 239
        FindDOSChar = Chr(166): Rem ο
    Case 240
        FindDOSChar = Chr(167): Rem π
    Case 241
        FindDOSChar = Chr(168): Rem ρ
    Case 242
        FindDOSChar = Chr(170): Rem ς
    Case 243
        FindDOSChar = Chr(169): Rem σ
    Case 244
        FindDOSChar = Chr(171): Rem τ
    Case 245
        FindDOSChar = Chr(172): Rem υ
    Case 246
        FindDOSChar = Chr(173): Rem φ
    Case 247
        FindDOSChar = Chr(174): Rem χ
    Case 248
        FindDOSChar = Chr(175): Rem ψ
    Case 249
        FindDOSChar = Chr(224): Rem ω
    
    Case 220
        FindDOSChar = Chr(152): Rem ά -> α
    Case 221
        FindDOSChar = Chr(156): Rem έ -> ε
    Case 222
        FindDOSChar = Chr(158): Rem ή -> η
    Case 223
        FindDOSChar = Chr(160): Rem ί -> ι
    Case 252
        FindDOSChar = Chr(166): Rem ό -> ο
    Case 253
        FindDOSChar = Chr(172): Rem ύ -> υ
    Case 254
        FindDOSChar = Chr(224): Rem ώ -> ω
    
    Case 162
        FindDOSChar = Chr(128): Rem 'Α -> Α
    Case 184
        FindDOSChar = Chr(132): Rem 'Ε -> Ε
    Case 185
        FindDOSChar = Chr(134): Rem 'Η -> Η
    Case 186
        FindDOSChar = Chr(136): Rem 'Ι -> Ι
    Case 188
        FindDOSChar = Chr(142): Rem 'Ο -> Ο
    Case 190
        FindDOSChar = Chr(147): Rem 'Υ -> Υ
    Case 191
        FindDOSChar = Chr(151): Rem 'Ω -> Ω
        
    Case 218
        FindDOSChar = Chr(136): Rem 'Ι -> Ι
    Case 219
        FindDOSChar = Chr(147): Rem 'Υ -> Υ
    Case 250
        FindDOSChar = Chr(160): Rem ϊ -> ι
    Case 251
        FindDOSChar = Chr(172): Rem ϋ -> υ
    
    Case Else
        FindDOSChar = strChar
    End Select

End Function

Function SelectRow(grdGrid, strKeyCode, lngRow, lngCol)

    'Βγαίνω
    If grdGrid.RowCount = 0 Or grdGrid.CellText(lngRow, lngCol) = "" Then SelectRow = 1: Exit Function
    
    'Μαρκάρω τη γραμμή
    With grdGrid
        If strKeyCode = 45 Or strKeyCode = 32 Then
            If .CellIcon(lngRow, "Selected") <= 0 Then
                SelectRow = 2
            Else
                SelectRow = 1
            End If
        End If
    End With

    'Ξεμαρκάρω τη γραμμή
    With grdGrid
        If strKeyCode = 46 Then
            SelectRow = 1: Exit Function
        End If
    End With

End Function

Function InitializeFields(ParamArray tmpFields())

    Dim bytLoop As Byte

    For bytLoop = 0 To UBound(tmpFields)
        If TypeOf tmpFields(bytLoop) Is newText Or TypeOf tmpFields(bytLoop) Is newDate Then tmpFields(bytLoop).text = format(Date, "dd/mm/yyyy")
        If TypeOf tmpFields(bytLoop) Is newInteger Then tmpFields(bytLoop).text = "0"
        If TypeOf tmpFields(bytLoop) Is newFloat Then tmpFields(bytLoop).text = "0,00"
        If TypeOf tmpFields(bytLoop) Is Label Then tmpFields(bytLoop).Caption = ""
        If TypeOf tmpFields(bytLoop) Is CheckBox Or TypeOf tmpFields(bytLoop) Is OptionButton Then tmpFields(bytLoop).Value = 1
    Next bytLoop
    
End Function

Function MoveToNextColumn(grdGrid As iGrid, lngRow, lngCol)

    On Error GoTo ErrTrap
    
    Do While True
        If lngCol + 1 <= grdGrid.colCount Then
            If grdGrid.ColTag(lngCol + 1) = "Y" Then
                grdGrid.SetCurCell lngRow, lngCol + 1
                Exit Function
            End If
        Else
            lngCol = 1
            Do While True
                grdGrid.SetCurCell lngRow + 1, lngCol
                If grdGrid.ColTag(lngCol) = "Y" Then
                    Exit Function
                End If
                lngCol = lngCol + 1
            Loop
        End If
        lngCol = lngCol + 1
    Loop
    
ErrTrap:
    If Err.Number = -2147220991 Then Exit Function

End Function

Function FillGridFromDB(SelectedDB, grdGrid, strTable, Fields, joins, criteriaString, sortColumn, ParamArray arguments())
    
    On Error GoTo ErrTrap
    
    Dim intLoop As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strSQL As String

    Dim rstTempRecordset As Recordset
    
    strPrinterName = ""
    FillGridFromDB = False
    
    strSQL = "SELECT " & IIf(Fields = "", "*", Fields) & " FROM " & strTable & " " & joins & IIf(criteriaString <> "", " WHERE " & criteriaString, "")
    
    Select Case SelectedDB
        Case "CommonDB"
            Set rstTempRecordset = CommonDB.OpenRecordset(strSQL)
        Case "PrintersDB"
            Set PrintersDB = DBEngine.OpenDataBase(App.Path + "\" + "Data" + "\" + "Printers.mdb", False, False)
            Set rstTempRecordset = PrintersDB.OpenRecordset(strSQL)
        Case "UsersDB"
            Set UsersDB = DBEngine.OpenDataBase(strPathName + "Users.mdb", False, False)
            Set rstTempRecordset = UsersDB.OpenRecordset(strSQL)
    End Select
    
    With grdGrid
        .Clear
        .Redraw = False
    End With
    
    Do Until rstTempRecordset.EOF
        grdGrid.AddRow
        intLoop = 0
        lngRow = grdGrid.RowCount
        For lngCol = 1 To UBound(arguments) + 1
            grdGrid.CellValue(lngRow, lngCol) = rstTempRecordset.Fields(arguments(intLoop))
            intLoop = intLoop + 1
        Next lngCol
        rstTempRecordset.MoveNext
    Loop
    
    grdGrid.Redraw = True
    
    If grdGrid.RowCount > 0 Then
        FillGridFromDB = True
        With grdGrid
            .Sort sortColumn
            .Enabled = True
        End With
    End If
    
    Exit Function
    
ErrTrap:
    FillGridFromDB = False
    DisplayErrorMessage True, Err.Description
    
    Exit Function
    
End Function

Public Function FormatDateAsFileName(myDate)

    If IsDate(myDate) Then
        FormatDateAsFileName = Right(myDate, 4) & "-" & Mid(myDate, 4, 2) & "-" & Left(myDate, 2)
    Else
        FormatDateAsFileName = myDate
    End If

End Function

Function CheckDateWithinLimits(WindowTitle, GivenDate, ClosedDate)

    'Αρχικές τιμές
    CheckDateWithinLimits = False
    
    'Κενή
    If GivenDate = "" Then
        If MyMsgBox(4, WindowTitle, strMessages(1), 1) Then
        End If
        Exit Function
    End If
    
    'Μεγαλύτερη της σημερινής
    If CDate(GivenDate) > Date Then
        If MyMsgBox(4, WindowTitle, strMessages(2), 1) Then
        End If
        Exit Function
    End If
    
    'Μικρότερη ή ίση της κλεισμένης περιόδου
    If CDate(GivenDate) <= ClosedDate Then
        If Not MyMsgBox(2, WindowTitle, strMessages(50), 2) Then
            Exit Function
        End If
    End If
    
    'Τελικές τιμές
    CheckDateWithinLimits = True
    
End Function

Private Function GetNewPID(username)

    Dim strPID As String
    
    strPID = username
    
    If (Len(strPID) > 20) Then
        strPID = Left$(strPID, 20)
    Else
        While (Len(strPID) < 4)
            strPID = strPID & "_"
        Wend
    End If
    
    GetNewPID = strPID
    
End Function

Public Function HashPassword(username, password)
    
    HashPassword = ToHexDump(CryptRC4(GetNewPID(username), password))

End Function

Function InitializeProgressBar(frmForm, lblTitle, tmpRecordset)
    
    On Error GoTo ErrTrap
    
    With frmForm
        If Not tmpRecordset.EOF Then
            frmForm.lblMaster.Caption = lblTitle
            frmForm.frmProgress.Visible = True
            frmForm.frmProgress.ZOrder 0
            frmForm.prgProgressBar.Value = 0
            frmForm.prgProgressBar.Min = 0
            If Not IsNumeric(tmpRecordset) Then
                tmpRecordset.MoveLast
                frmForm.prgProgressBar.Max = tmpRecordset.RecordCount
                tmpRecordset.MoveFirst
            Else
                frmForm.prgProgressBar.Max = tmpRecordset
            End If
            frmForm.Refresh
        End If
    End With
    
    Exit Function
    
ErrTrap:
    
    If Err.Number = 424 Then
        Resume Next
    End If

End Function

Function PrintHeadings(tmpColumns, tmpPageNo, tmpReportTitle, tmpReportSubTitle1, tmpReportSubTitle2, tmpTopMargin)

    Dim intLeft As Integer
    Dim intPageLength As Integer
    Dim intTopMargin As Integer
    
    intPageLength = 6 + Len(tmpPageNo)
    
    For intTopMargin = 1 To intTopMargin - 1
        Print #1, ""
    Next intTopMargin

    Print #1, arrCompanyData(7); Tab(tmpColumns - intPageLength); "ΣΕΛΙΔΑ " & tmpPageNo
    Print #1, arrCompanyData(8)
    Print #1, arrCompanyData(9)
    Print #1, arrCompanyData(10)
    
    Print #1, ""
    
    intLeft = (tmpColumns / 2) - (Len(tmpReportTitle) / 2)
    Print #1, Space(intLeft) & tmpReportTitle
    intLeft = (tmpColumns / 2) - (Len(tmpReportSubTitle1) / 2)
    If tmpReportSubTitle1 <> "" Then Print #1, Space(intLeft) & tmpReportSubTitle1
    intLeft = (tmpColumns / 2) - (Len(tmpReportSubTitle2) / 2)
    If tmpReportSubTitle2 <> "" Then Print #1, Space(intLeft); tmpReportSubTitle2
    
    Print #1, ""
    
End Function

Function PrintColumnHeadings(ParamArray columns() As Variant)

    Dim bytLoop As Byte
    
    For bytLoop = 0 To UBound(columns) - 1 Step 2
        Print #1, Tab(columns(bytLoop)); columns(bytLoop + 1);
    Next bytLoop
    
    Print #1, ""
    
End Function

Function SimpleSeek(Table, Index, ParamArray Indexes() As Variant)

    On Error GoTo ErrTrap
    
    Dim intLoop As Integer
    Dim intInnerLoop As Integer
    Dim strField()
    Dim intUpper As Integer
    Dim intArrayindex As Integer
    Dim strNewField As String
    Dim rsTable As Recordset
    
    SimpleSeek = False
    
    Set rsTable = CommonDB.OpenRecordset(Table)

    With rsTable
        .Index = Index
        If UBound(Indexes) = 0 Then .Seek "=", Indexes(0) Else .Seek "=", Indexes(0), Indexes(1)
        If .NoMatch Then SimpleSeek = True 'Αν η εγγραφή δεν βρεθεί, μπορώ να την διαγράψω
        .Close
    End With
    
    Exit Function
    
ErrTrap:
    SimpleSeek = False
    DisplayErrorMessage True, Err.Description

End Function

Function UpdateProgressBar(frmForm)
    
    frmForm.prgProgressBar.Value = frmForm.prgProgressBar.Value + 1
       
End Function

Function CheckForMatch(DBToUse, myFieldValue, myTable, myFieldLookup, myFieldType, myShowInList, myOrder, Optional checkWholeField As Boolean) As Recordset

    On Error GoTo ErrTrap
    
    'Local μεταβλητές
    Dim lngCol As Long
    Dim lngRow As Long
    Dim bytLoop As Byte
    Dim bytGroupStart As Byte
    Dim bytArrayIndex As Byte
    Dim arrFirstElements(), arrSecondElements(), arrThirdElements(), arrFourthElements()
    
    Dim TempFields As typTableData
    Dim rstTempRecordset As Recordset
    
    'Επιλογή βάσης
    If DBToUse = "PrintersDB" Then
        Set PrintersDB = DBEngine.OpenDataBase(App.Path + "\" + "Data" + "\" + "Printers.mdb", False, False)
        Set TempQuery = PrintersDB.CreateQueryDef("")
    Else
        Set TempQuery = CommonDB.CreateQueryDef("")
    End If
    
    Do While True
        'Αν δεν έχω δώσει τίποτα
        If Len(myFieldValue) = 0 Then
            If myShowInList <> "" Then
                TempQuery.SQL = "SELECT * FROM " & myTable & " WHERE ShowInList = " & myShowInList & " ORDER BY " & myOrder
            Else
                TempQuery.SQL = "SELECT * FROM " & myTable & " ORDER BY 1"
            End If
            Exit Do
        End If
        'Αν έχω δώσει αριθμό
        If myFieldType = "Numeric" Then
            TempQuery.SQL = "PARAMETERS lngCode Long; " _
            & "SELECT * FROM " & myTable & " WHERE " _
            & "[" & myFieldLookup & "] = " & Val(myFieldValue) & " "
            If myShowInList <> 0 Then TempQuery.SQL = TempQuery.SQL & " AND ShowInList = " & myShowInList & " ORDER BY " & myOrder
            TempQuery![lngCode] = Val(myFieldValue)
            Exit Do
        'Αν έχω δώσει κείμενο
        Else
            If Left(myFieldValue, 1) <> "*" Then
                TempQuery.SQL = "PARAMETERS strDescription String; " _
                & "SELECT * FROM " & myTable & " WHERE " _
                & IIf(Not checkWholeField, "Left(" & myFieldLookup & ",Len(strDescription)) = " & "'" & myFieldValue & "' ", "" & myFieldLookup & " = '" & myFieldValue & "'")
                If myShowInList <> 0 Then TempQuery.SQL = TempQuery.SQL & " AND ShowInList = " & myShowInList & " ORDER BY " & myOrder
                TempQuery![strDescription] = myFieldValue
            Else
                If Len(myFieldValue) > 1 Then
                    TempQuery.SQL = "PARAMETERS strDescription String; " _
                    & "SELECT * FROM " & myTable & " WHERE " _
                    & "InStr([" & myFieldLookup & "], " & "'" & Right(myFieldValue, Len(myFieldValue) - 1) & "'" & ") "
                    If myShowInList <> 99 Then
                        TempQuery.SQL = TempQuery.SQL & " AND ShowInList = " & myShowInList & " ORDER BY " & myOrder
                        TempQuery![strDescription] = Right(myFieldValue, Len(myFieldValue) - 1)
                    End If
                Else
                    TempQuery.SQL = "PARAMETERS strDescription String; " _
                    & "SELECT * FROM " & myTable
                    If myShowInList <> 99 Then
                        TempQuery.SQL = TempQuery.SQL & " WHERE ShowInList = " & myShowInList & " ORDER BY " & myOrder
                        TempQuery![strDescription] = myFieldValue
                    End If
                End If
            End If
            Exit Do
        End If
    Loop
    
    Set rstTempRecordset = TempQuery.OpenRecordset()
    
    Set CheckForMatch = rstTempRecordset
    
    Exit Function
    
ErrTrap:
    If Err.Number <> 3075 Then
        DisplayErrorMessage True, Err.Description
    End If
    
End Function

Function SetUpGrid(myIconList As vbalImageList, ParamArray myGrid() As Variant)
    
    Dim intLoop As Integer
    
    For intLoop = 0 To UBound(myGrid)
        With myGrid(intLoop)
            .Editable = False
            .DefaultRowHeight = 23
            .RowMode = True
            .ScrollBarStyle = 2
            .Top = .Top - 6
            With .Font
                .Name = "Ubuntu Condensed"
                .Size = 12
                .Bold = False
            End With
            With .Header
                .Flat = True
                .Buttons = False
                .BackColor = GetSetting(appName:=strAppTitle, Section:="Appearance", Key:="Grid Header BackColor")
                .ForeColor = GetSetting(appName:=strAppTitle, Section:="Appearance", Key:="Grid Header ForeColor")
                .SortInfoStyle = igSortInfoNone
                With .Font
                    .Name = "Ubuntu Condensed"
                    .Size = 10
                End With
            End With
            .ImageList = myIconList
        End With
    Next intLoop

End Function

Function CaptureNumbers(strString, tmpRow, tmpCol, tmpKeyAscii, blnDecimals)

    If (tmpKeyAscii = 46 Or tmpKeyAscii = 44) And blnDecimals Then
        If InStr(strString, ".") Or InStr(strString, ",") Then
            tmpKeyAscii = 0
        Else
            tmpKeyAscii = 44
            Exit Function
        End If
        Exit Function
    End If
    
    If (tmpKeyAscii < 48 Or tmpKeyAscii > 58) And tmpKeyAscii <> 8 And tmpKeyAscii <> 13 Then
        tmpKeyAscii = 0
    End If

End Function

Function MainSeekRecord(SelectedDB, Table, IndexField, CodeToSeek, DisplayNotFoundMessage, ParamArray Fields())

    On Error GoTo ErrTrap
    
    Dim bytLoop As Byte
    Dim rsTable As Recordset
    
    Select Case SelectedDB
        Case "CommonDB"
            Set rsTable = CommonDB.OpenRecordset(Table)
        Case "PrintersDB"
            Set rsTable = PrintersDB.OpenRecordset(Table)
        Case "UsersDB"
            Set rsTable = UsersDB.OpenRecordset(Table)
    End Select
    
    MainSeekRecord = True
    
    With rsTable
        .Index = IndexField
        .Seek "=", CodeToSeek
        If Not .NoMatch Then
            For bytLoop = 0 To UBound(Fields)
                If TypeOf Fields(bytLoop) Is TextBox Or TypeOf Fields(bytLoop) Is newText Then
                    Fields(bytLoop).text = IIf(Not IsNull(rsTable.Fields(bytLoop)), rsTable.Fields(bytLoop), "")
                End If
                If TypeOf Fields(bytLoop) Is newFloat Then
                    Fields(bytLoop).text = format(rsTable.Fields(bytLoop), "#,##0.00")
                End If
                If TypeOf Fields(bytLoop) Is newInteger Then
                    Fields(bytLoop).text = format(rsTable.Fields(bytLoop), "#,##0")
                End If
                If TypeOf Fields(bytLoop) Is Label Then
                    Fields(bytLoop).Caption = rsTable.Fields(bytLoop)
                End If
                If TypeOf Fields(bytLoop) Is CheckBox Then
                    Fields(bytLoop).Value = IIf(rsTable.Fields(bytLoop), 1, 0)
                End If
                If TypeOf Fields(bytLoop) Is OptionButton Then
                    Fields(bytLoop).Value = IIf(rsTable.Fields(bytLoop), 1, 0)
                End If
                If TypeOf Fields(bytLoop) Is newDate Then
                    Fields(bytLoop).text = format(rsTable.Fields(bytLoop), "dd/mm/yyyy")
                End If
            Next bytLoop
        Else
            If DisplayNotFoundMessage Then
                If MyMsgBox(4, strAppTitle, strMessages(17), 1) Then
                End If
                MainSeekRecord = False
            End If
        End If
        .Close
    End With
    
    Exit Function
    
ErrTrap:
    MainSeekRecord = False
    DisplayErrorMessage True, Err.Description
    
    Exit Function

End Function

Function HighlightRow(grdGrid As iGrid, lngColumn, strID, blnRowMode)

    Dim lngRow As Long
    
    If strID <> "" Then
        With grdGrid
            For lngRow = 1 To .RowCount
                If (.CellText(lngRow, lngColumn) = strID) Then
                    .EnsureVisibleRow lngRow
                    .SetCurCell lngRow, lngColumn
                    .RowMode = blnRowMode
                    .SetFocus
                End If
            Next lngRow
        End With
    End If
    
    If strID = "" Then
        With grdGrid
            .EnsureVisibleRow 1
            .SetCurCell 1, lngColumn
            .RowMode = blnRowMode
            .SetFocus
        End With
    End If

End Function

Function MyMsgBox(intPictureIndex, txtTitle, txtLine, intNoOfButtons, Optional errorDescription = "")

    With CommonMessages
        .frmButtonFrame(1).Visible = False
        .frmButtonFrame(2).Visible = False
        .imgImage.Picture = .lslIcons.ItemPicture(intPictureIndex)
        .imgImage.ToolTipText = errorDescription
        .lblTitle = txtTitle
        .lblLine = txtLine
        .frmButtonFrame(intNoOfButtons).Visible = True
        .Show 1
        If .cmdButton(0).Tag = "Pressed" Then
            MyMsgBox = True
            Exit Function
        Else
            MyMsgBox = False
            Exit Function
        End If
        If .cmdButton(2).Tag = "Pressed" Then
            MyMsgBox = True
        End If
    End With
    
End Function

Function DisplayErrorMessage(DisplayMessage, errorDescription, Optional progress As Frame, Optional grid As iGrid, Optional CloseThisConnection As Boolean = True)

    If DisplayMessage Then
        If Not progress Is Nothing Then progress.Visible = False
        If Not grid Is Nothing Then grid.Redraw = True
        If MyMsgBox(4, strAppTitle, strMessages(26), 1, errorDescription) Then
        End If
    End If
    
    UpdateLogFile errorDescription

End Function

Function DisplayIndex(tmpRecordset, blnShowList, blnIncludeOneRecordCount, strTitle, tmpGroupElements, ParamArray tmpArguments()) As typTableData

    On Error GoTo TrapError
    
    Dim bytLoop As Byte
    
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngActiveColumn As Long
    
    Dim TempFields As typTableData
    
    If Not tmpRecordset.EOF Then
        tmpRecordset.MoveFirst
        GoSub InitializeGrid
        While tmpRecordset.EOF = False
            With CommonIndex.grdGrid
                .AddRow
                bytLoop = 0
                lngRow = .RowCount
                For lngCol = 1 To tmpGroupElements
                    .CellValue(lngRow, lngCol) = tmpRecordset.Fields(tmpArguments(bytLoop))
                    bytLoop = bytLoop + 1
                Next lngCol
            End With
            tmpRecordset.MoveNext
        Wend
        
        If blnIncludeOneRecordCount Or CommonIndex.grdGrid.RowCount > 1 Then
            If blnShowList Then
                CommonIndex.grdGrid.Redraw = True
                If CommonIndex.grdGrid.HScrollBar.Visible Then
                    Do Until Not CommonIndex.grdGrid.HScrollBar.Visible
                        CommonIndex.grdGrid.Width = CommonIndex.grdGrid.Width + 90
                    Loop
                    GoSub ResizeForm
                    GoSub PositionInactiveRecordsCheckBox
                    GoSub DisplayButtonForLedger
                End If
                With CommonIndex
                    .lblTitle.Caption = strTitle
                    .chkShowInactiveRecords.Visible = False
                    
                    If tmpArguments(UBound(tmpArguments)) = "Persons" Then
                        lngActiveColumn = tmpGroupElements
                        .chkShowInactiveRecords.Visible = True
                        GoSub DisplayOnlyActiveItems
                    End If
                    
                    If tmpArguments(UBound(tmpArguments)) = "Items" Then
                        lngActiveColumn = 10
                        GoSub DisplayOnlyWithCategoryCheckBalanceIsTrue
                        GoSub DisplayOnlyActiveItems
                        GoSub PaintWithAlternateColor
                    End If
                    
                    .grdGrid.Enabled = True
                    .grdGrid.Redraw = True
                    
                    For lngRow = 1 To .grdGrid.RowCount
                        If .grdGrid.RowVisible(lngRow) Then
                            .grdGrid.SetCurCell lngRow, 1
                            Exit For
                        End If
                    Next lngRow
                    
                    .Show 1
                End With
            End If
        Else
            If tmpArguments(UBound(tmpArguments)) = "Persons" Then
                lngActiveColumn = tmpGroupElements
            End If

            CommonIndex.grdGrid.CurRow = 1
        End If
    End If
    
    With CommonIndex
        TempFields.strCode = .grdGrid.CellValue(.grdGrid.CurRow, 1)
        TempFields.strOneField = .grdGrid.CellValue(.grdGrid.CurRow, 2)
        TempFields.strTwoField = .grdGrid.CellValue(.grdGrid.CurRow, 3)
        TempFields.strThreeField = .grdGrid.CellValue(.grdGrid.CurRow, 4)
        TempFields.strFourField = .grdGrid.CellValue(.grdGrid.CurRow, 5)
        TempFields.strFiveField = .grdGrid.CellValue(.grdGrid.CurRow, 6)
        TempFields.strSixField = .grdGrid.CellValue(.grdGrid.CurRow, 7)
        TempFields.strSevenField = .grdGrid.CellValue(.grdGrid.CurRow, 8)
        TempFields.strEightField = .grdGrid.CellValue(.grdGrid.CurRow, 9)
        TempFields.strNineField = .grdGrid.CellValue(.grdGrid.CurRow, 10)
        TempFields.strTenField = .grdGrid.CellValue(.grdGrid.CurRow, 11)
        TempFields.strElevenField = .grdGrid.CellValue(.grdGrid.CurRow, 12)
        TempFields.strTwelveField = .grdGrid.CellValue(.grdGrid.CurRow, 13)
        TempFields.strThirteenField = .grdGrid.CellValue(.grdGrid.CurRow, 14)
        TempFields.strFourteenField = .grdGrid.CellValue(.grdGrid.CurRow, 15)
        TempFields.strFifteenField = .grdGrid.CellValue(.grdGrid.CurRow, 16)
        TempFields.strSixteenField = .grdGrid.CellValue(.grdGrid.CurRow, 17)
        TempFields.strSeventeenField = .grdGrid.CellValue(.grdGrid.CurRow, 18)
        
    End With
    
    DisplayIndex = TempFields
    
    Unload CommonIndex
    
    Exit Function
    
TrapError:
    If Err.Number = 3021 Or Err.Number = 91 Or Err.Number = -2147220991 Or Err.Number = 3265 Or Err.Number = 3075 Then
        DisplayIndex = TempFields
        Unload CommonIndex
        Exit Function
    Else
    
    End If
    If Err.Number = 94 Then
        Resume Next
    End If

InitializeGrid:
    
    ReDim arrFirstElements(1)
    ReDim arrSecondElements(1)
    ReDim arrThirdElements(1)
    ReDim arrFourthElements(1)
    
    Dim bytGroupStart As Byte
    Dim bytArrayIndex As Byte
    
    For bytLoop = 0 To UBound(tmpArguments) + 1
        'Περιεχόμενο
        bytGroupStart = tmpGroupElements
        bytArrayIndex = 1
        While bytLoop < tmpGroupElements
            ReDim Preserve arrFirstElements(UBound(arrFirstElements))
            arrFirstElements(bytArrayIndex) = tmpRecordset(tmpArguments(bytLoop))
            bytLoop = bytLoop + 1
        Wend
        'Τίτλος Στήλης
        bytGroupStart = tmpGroupElements + bytGroupStart
        bytArrayIndex = 1
        While bytLoop < bytGroupStart
            ReDim Preserve arrSecondElements(UBound(arrSecondElements) + 1)
            arrSecondElements(bytArrayIndex) = tmpArguments(bytLoop)
            bytArrayIndex = bytArrayIndex + 1
            bytLoop = bytLoop + 1
        Wend
        'Πλάτος Στηλών
        bytGroupStart = tmpGroupElements + bytGroupStart
        bytArrayIndex = 1
        While bytLoop < bytGroupStart
            ReDim Preserve arrThirdElements(UBound(arrThirdElements) + 1)
            arrThirdElements(bytArrayIndex) = tmpArguments(bytLoop)
            bytArrayIndex = bytArrayIndex + 1
            bytLoop = bytLoop + 1
        Wend
        'Στοίχιση Στηλών
        bytGroupStart = tmpGroupElements + bytGroupStart
        bytArrayIndex = 1
        While bytLoop < bytGroupStart
            ReDim Preserve arrFourthElements(UBound(arrFourthElements) + 1)
            arrFourthElements(bytArrayIndex) = tmpArguments(bytLoop)
            bytArrayIndex = bytArrayIndex + 1
            bytLoop = bytLoop + 1
        Wend
    Next bytLoop
    
    'Προσθέτω στήλες - τίτλους - πλάτη
    CommonIndex.grdGrid.Width = 0
    For bytLoop = 1 To tmpGroupElements
        CommonIndex.grdGrid.AddCol.eTextFlags = arrFourthElements(bytLoop)
        CommonIndex.grdGrid.ColHeaderText(bytLoop) = arrSecondElements(bytLoop)
        CommonIndex.grdGrid.ColWidth(bytLoop) = 7 * (arrThirdElements(bytLoop) + 1)
        If arrThirdElements(bytLoop) = 0 Then CommonIndex.grdGrid.ColVisible(bytLoop) = False
        CommonIndex.grdGrid.ColHeaderTextFlags(bytLoop) = 1
    Next bytLoop
    
    With CommonIndex.grdGrid
        .Header.Flat = True
        .Header.Height = 25
    End With
        
    Return
    
DisplayOnlyWithCategoryCheckBalanceIsTrue:
    For lngRow = 1 To CommonIndex.grdGrid.RowCount
         If CommonIndex.grdGrid.CellValue(lngRow, CommonIndex.grdGrid.colCount) = "0" Then
            CommonIndex.grdGrid.CellValue(lngRow, 9) = ""
        End If
    Next lngRow

    Return


DisplayOnlyActiveItems:
    
    For lngRow = 1 To CommonIndex.grdGrid.RowCount
        If CommonIndex.grdGrid.RowVisible(lngRow) <> CommonIndex.grdGrid.CellValue(lngRow, lngActiveColumn) Then
            For lngCol = 1 To CommonIndex.grdGrid.colCount
                CommonIndex.grdGrid.CellFont(lngRow, lngCol).Italic = True
                CommonIndex.grdGrid.CellForeColor(lngRow, lngCol) = &HC0C0C0
                CommonIndex.grdGrid.RowVisible(lngRow) = False
            Next lngCol
        End If
        
    Next lngRow
    
    Return
    
PaintWithAlternateColor:

    Dim lngBackColor As Long
    Dim lngForeColor As Long
    Dim strOldManufacturer As String
    Dim blnFirstTimeManufacturer As Boolean

    With CommonIndex.grdGrid
        
        For lngRow = 1 To .RowCount
            If .RowVisible(lngRow) Then
                lngBackColor = &HC8C8FF
                lngForeColor = vbBlack
                blnFirstTimeManufacturer = True
                strOldManufacturer = .CellValue(lngRow, 6)
                Exit For
            End If
        Next lngRow
        
        For lngRow = lngRow To .RowCount
            If .RowVisible(lngRow) Then
                If strOldManufacturer = .CellValue(lngRow, 6) Then
                    If .CellValue(lngRow, 6) <> strOldManufacturer Then
                        lngBackColor = IIf(lngBackColor = -1, &HC8C8FF, -1)
                        lngForeColor = IIf(lngForeColor = -1, vbBlack, -1)
                        For lngCol = 1 To .colCount
                            .CellForeColor(lngRow, lngCol) = lngForeColor
                            .CellBackColor(lngRow, lngCol) = lngBackColor
                        Next
                    Else
                        For lngCol = 1 To .colCount
                            .CellForeColor(lngRow, lngCol) = lngForeColor
                            .CellBackColor(lngRow, lngCol) = lngBackColor
                        Next
                    End If
                Else
                    strOldManufacturer = .CellValue(lngRow, 6)
                    lngBackColor = IIf(lngBackColor = -1, &HC8C8FF, -1)
                    lngForeColor = IIf(lngForeColor = -1, vbBlack, -1)
                    For lngCol = 1 To .colCount
                        .CellForeColor(lngRow, lngCol) = lngForeColor
                        .CellBackColor(lngRow, lngCol) = lngBackColor
                    Next
                End If
            End If
        Next
        
    End With

Return

ResizeForm:
    
    With CommonIndex
        .shpShape.Width = .grdGrid.Width + 160
        .Width = .shpShape.Width + 470
        .frmButtonFrame.Left = (.Width / 2) - (.frmButtonFrame.Width / 2)
    End With

    Return
    
PositionInactiveRecordsCheckBox:

    CommonIndex.chkShowInactiveRecords.Left = CommonIndex.Width - CommonIndex.chkShowInactiveRecords.Width - 320

    Return
    
DisplayButtonForLedger:

    If tmpArguments(UBound(tmpArguments)) = "Items" Then
        CommonIndex.frmButtonFrame.Width = 4365
        CommonIndex.cmdButton(2).Left = 2925
        CommonIndex.cmdButton(1).Enabled = True
    Else
        CommonIndex.frmButtonFrame.Width = 2940
        CommonIndex.cmdButton(2).Left = 1500
        CommonIndex.cmdButton(1).Enabled = False
    End If
    
    CommonIndex.frmButtonFrame.Left = CommonIndex.Width / 2 - CommonIndex.frmButtonFrame.Width / 2

    Return
    
End Function

Function PrintRecords(myForm As Form, myWhatToDo, myDisplayCompletionMessage, myInvoiceOrReport, Optional myPrinterCodeID, Optional myInvoiceTrnID)

    ' strCode = Ονομα
    ' strOneField = Φιλική ονομασία
    ' strTwoField = ID Τύπου
    ' strThreeField = Ονομα γραμματοσειράς
    ' strFourField = Μέγεθος γραμματοσειράς
    
    ' strFiveField = String σήμανσης
    ' strSixField = Υψος παραστατικού
    ' strSevenField = Γραμμές παραστατικού
    ' strEightField = Επάνω περιθώριο παραστατικού
    
    ' strNineField = Κωδ. χαρτιού
    ' strTenField = ID προσανατολισμού
    
    ' strElevenField = Υψος αναφορών
    ' strTwelveField = Γραμμές αναφορών
    ' strThirteenField = Επάνω περιθώριο αναφορών
    ' strFourteenField = Αριστερο περιθώριο αναφορών

    On Error GoTo ErrTrap
    
    Dim strFileName As String
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    If myWhatToDo = "Print" Then
        'Αν τυπώνω το τιμολογιο, βρισκω τον εκτυπωτη
        If Not IsMissing(myPrinterCodeID) Then
            Set tmpRecordset = CheckForMatch("PrintersDB", myPrinterCodeID, "Printers", "PrinterID", "Numeric", 1, 1)
        Else
            Set tmpRecordset = CheckForMatch("PrintersDB", 1, "Printers", myInvoiceOrReport, "Numeric", 1, 1)
        End If
        If tmpRecordset.RecordCount = 0 Then DisplayMessage IIf(myInvoiceOrReport = "Report", 5, 61), 4, 1, "": Exit Function
        tmpTableData = DisplayIndex(tmpRecordset, True, False, "Εκτυπωτές", 15, 1, 2, 3, 4, 5, 8, 9, 10, 11, 13, 14, 15, 16, 17, 18, "Ονομα", "Φιλική ονομασία", "ID Τύπου", "Ονομα γραμματοσειράς", "Μέγεθος γραμματοσειράς", "String σήμανσης", "Υψος παραστατικού", "Γραμμές παραστατικού", "Επάνω περιθώριο παραστατικού", "Κωδ. χαρτιού", "ID προσανατολισμού", "Υψος αναφορών", "Γραμμές αναφορών", "Επάνω περιθώριο αναφορών", "Αριστερο περιθώριο αναφορών", 0, 40, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        With tmpTableData
            If .strCode <> "" Then
                If PrinterExists(.strCode) Then
                    If myInvoiceOrReport = "PrinterPrintsInvoicesID" Then strFileName = myForm.CreateUnicodeFile(.strTwoField, .strFiveField, Val(.strSixField), Val(.strSevenField), Val(.strEightField), 0) 'Παραστατικά - Πάντα φτιάχνω το unicode (ID τύπου εκτυπωτή, string σήμανσης = 5, ύψος = 6, αναλυτικές γραμμές = 7, επάνω περιθώριο = 8, αριστερό περιθώριο = 9)
                    If myInvoiceOrReport = "PrinterPrintsReportsID" Then strFileName = myForm.CreateUnicodeFile(.strTwoField, "", Val(.strElevenField), Val(.strTwelveField), Val(.strThirteenField), Val(.strFourteenField)) 'Αναφορές - Πάντα φτιάχνω το unicode (string σήμανσης = "", ύψος = 11, αναλυτικές γραμμές = 12, επάνω περιθώριο = 13, αριστερό περιθώριο = 14)
                    If .strTwoField = "1" Then strFileName = ConvertToAsciiFile(strFileName, strAsciiFile)  'Αν ο εκτυπωτής είναι dot matrix, μετατρέπω το unicode σε ascii
                    If .strTwoField <> "1" Then 'Αν ο εκτυπωτής ΔΕΝ είναι dot matrix
                        CreatePDF Val(tmpTableData.strNineField), Val(tmpTableData.strTenField), Val(tmpTableData.strThirteenField), Val(tmpTableData.strFourteenField), myForm.lblTitle.Caption, tmpTableData.strThreeField, Val(tmpTableData.strFourField), myInvoiceOrReport '
                        If myInvoiceOrReport = "PrinterPrintsInvoicesID" Then PrintInvoiceToLaser myInvoiceTrnID, tmpTableData.strCode
                        If myInvoiceOrReport = "PrinterPrintsReportsID" And Not blnPreviewReports Then PrintPDF .strCode
                    End If
                Else
                    If MyMsgBox(4, strAppTitle, strMessages(18), 1) Then
                    End If
                    Exit Function
                End If
            End If
        End With
    End If
    
    If myWhatToDo = "CreatePDF" Then
        If ExportToPDF(myForm.CreateUnicodeFile("", "", GetSetting(strAppTitle, "Settings", "ExportReportHeight"), GetSetting(strAppTitle, "Settings", "ExportReportDetailLines"), GetSetting(strAppTitle, "Settings", "ExportReportTopMargin"), GetSetting(strAppTitle, "Settings", "ExportReportLeftMargin"))) Then
            If myDisplayCompletionMessage Then
                DisplayMessage 10, 1, 1, ""
                Exit Function
            End If
        End If
    End If
    
    Exit Function
    
ErrTrap:
    Exit Function
    
End Function

Function DisplayMessageRecordsNotFound()

    If MyMsgBox(1, strAppTitle, strMessages(8), 1) Then
    End If

End Function

Function MainSaveRecord(SelectedDB, Table, Status, FormTitle, IndexField, CodeToSeek, ParamArray Fields() As Variant)

    On Error GoTo ErrTrap
    
    Dim lngFieldNo As Long
    Dim rsTable As Recordset
    
    Select Case SelectedDB
        Case "CommonDB"
            Set rsTable = CommonDB.OpenRecordset(Table)
        Case "PrintersDB"
            Set rsTable = PrintersDB.OpenRecordset(Table)
        Case "UsersDB"
            Set rsTable = UsersDB.OpenRecordset(Table)
    End Select
    
    With rsTable
        .Index = IndexField
        If Status Then
            .AddNew
        Else
            .Seek "=", CodeToSeek
            If Not .NoMatch Then
                .Edit
            Else
                If MyMsgBox(4, FormTitle, strMessages(17), 1) Then
                End If
                Exit Function
            End If
        End If
        For lngFieldNo = 0 To UBound(Fields)
            Debug.Print .Fields(lngFieldNo + 1).Name & " " & Fields(lngFieldNo)
            .Fields(lngFieldNo + 1).Value = Trim(Fields(lngFieldNo))
        Next
        .Update
        If Status Then
            .MoveLast
        End If
        MainSaveRecord = .Fields(0).Value
        .Close
    End With
    
    Exit Function
    
ErrTrap:
    MainSaveRecord = False
    DisplayErrorMessage True, Err.Description
    
End Function

Function UpdateButtons(formName As Form, Max, ParamArray Buttons() As Variant)
    
    Dim intLoop As Integer
    
    For intLoop = 0 To UBound(Buttons)
        formName.cmdButton(intLoop).MousePointer = vbCrosshair
        formName.cmdButton(intLoop).Enabled = IIf(Buttons(intLoop) = 0, False, True)
    Next intLoop
    
End Function

Function UpdateRecordCount(myLabel As Label, myRecordCount)

    myLabel.Caption = "Βρέθηκαν " & myRecordCount & " εγγραφές"

End Function

Function CustomUpperCase(strString)

    Dim intLoop As Integer
    Dim strNewString As String
    
    For intLoop = 1 To Len(strString)
        strNewString = strNewString & ConvertCharacterToUpperCase(Mid(strString, intLoop, 1))
    Next intLoop
    
    CustomUpperCase = strNewString

End Function

Function ColorizeCells(myGrid As iGrid, myRow, ParamArray myCols() As Variant)

    Dim lngCol As Long
    Dim intLoop As Integer
    
    For intLoop = 0 To UBound(myCols())
        For lngCol = 1 To myGrid.colCount
            If myGrid.ColKey(lngCol) = myCols(intLoop) Then
                myGrid.CellForeColor(myRow, lngCol) = IIf(Left(myGrid.CellValue(myRow, lngCol), 1) <> "-", -1, &H8080FF)
                Exit For
            End If
        Next lngCol
    Next intLoop

End Function

Function ValidateInput(KeyAscii)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Sendkeys "{TAB}"
    End If

End Function

Public Function Sendkeys(text As Variant, Optional wait As Boolean = False)
   
    Dim WshShell As Object
   
    Set WshShell = CreateObject("wscript.shell")
   
    WshShell.Sendkeys CStr(text), wait
   
    Set WshShell = Nothing
   
End Function

Function ClearFields(ParamArray tmpFields())

    Dim bytLoop As Byte
    
    For bytLoop = 0 To UBound(tmpFields)
        If TypeOf tmpFields(bytLoop) Is Field Then
            tmpFields(bytLoop).text = ""
        End If
        If TypeOf tmpFields(bytLoop) Is TextBox Or TypeOf tmpFields(bytLoop) Is newText Or TypeOf tmpFields(bytLoop) Is newInteger Or TypeOf tmpFields(bytLoop) Is newFloat Or TypeOf tmpFields(bytLoop) Is newDate Then
            tmpFields(bytLoop).text = ""
        End If
        If TypeOf tmpFields(bytLoop) Is Label Then
            tmpFields(bytLoop).Caption = ""
        End If
        If TypeOf tmpFields(bytLoop) Is CheckBox Or TypeOf tmpFields(bytLoop) Is OptionButton Then
            tmpFields(bytLoop).Value = 0
        End If
        If TypeOf tmpFields(bytLoop) Is Image Then
            tmpFields(bytLoop).Picture = LoadPicture(strImageDirectory & "Κενό16x16.ico")
        End If
        If TypeOf tmpFields(bytLoop) Is ComboBox Then
            tmpFields(bytLoop).ListIndex = -1
        End If
        If TypeOf tmpFields(bytLoop) Is iGrid Then
            tmpFields(bytLoop).Clear
            tmpFields(bytLoop).GridLines = igGridLinesNone
        End If
        If TypeOf tmpFields(bytLoop) Is Frame Then
            tmpFields(bytLoop).Visible = False
        End If
    Next bytLoop

End Function

Function DisableFields(ParamArray tmpFields())

    Dim bytLoop As Byte
    
    For bytLoop = 0 To UBound(tmpFields)
        tmpFields(bytLoop).Enabled = False
    Next bytLoop

End Function

Function ExportToPDF(fileName As String)

    On Error GoTo ErrTrap
    
    Dim pdf As New ARExportPDF

    ClearVariables strPrinterData(1), strPrinterData(2), strPrinterData(3), strPrinterData(4), strPrinterData(5)

    With rptReport
        .Restart
        .oneLongField.Width = 10100
        .oneLongField.Font.Name = "Input"
        .oneLongField.Font.Size = 7
        .Run False
        pdf.AcrobatVersion = 2
        pdf.SemiDelimitedNeverEmbedFonts = ""
        pdf.fileName = Replace(fileName, "/", "-")
        pdf.fileName = Replace(pdf.fileName, "[", "")
        pdf.fileName = Replace(pdf.fileName, "]", "")
        pdf.fileName = Replace(pdf.fileName, "  ", " ")
        pdf.fileName = Replace(pdf.fileName, ".txt", ".pdf")
        pdf.Export .Pages
    End With
    
    ExportToPDF = True
    
    Exit Function
    
ErrTrap:
    ExportToPDF = False
    DisplayErrorMessage True, Err.Description

End Function

Function MainDeleteRecord(SelectedDB, Table, FormTitle, IndexField, CodeToSeek, AskConfirmation)

    On Error GoTo ErrTrap
    
    Dim rsTable As Recordset
    
    Select Case SelectedDB
        Case "CommonDB"
            Set rsTable = CommonDB.OpenRecordset(Table)
        Case "PrintersDB"
            Set rsTable = PrintersDB.OpenRecordset(Table)
        Case "UsersDB"
            Set rsTable = UsersDB.OpenRecordset(Table)
    End Select

    With rsTable
        .Index = IndexField
        .Seek "=", CodeToSeek
        If Not .NoMatch Then
            If AskConfirmation = "False" Then
                .Delete
                .Close
                MainDeleteRecord = True
                Exit Function
            End If
            If MyMsgBox(3, FormTitle, strMessages(4), 2) Then
                .Delete
                .Close
                MainDeleteRecord = True
            Else
                .Close
                MainDeleteRecord = False
            End If
        Else
            If MyMsgBox(4, FormTitle, strMessages(17), 1) Then
            End If
        End If
    End With
    
    Exit Function
    
ErrTrap:
    MainDeleteRecord = False
    DisplayErrorMessage True, Err.Description
    
End Function

Function EnableFields(ParamArray tmpFields())
    
    Dim bytLoop As Byte
    
    For bytLoop = 0 To UBound(tmpFields)
        tmpFields(bytLoop).Enabled = True
        If TypeOf tmpFields(bytLoop) Is iGrid Then
            tmpFields(bytLoop).Editable = True
        End If
    Next bytLoop

End Function

Function EditableFields(ParamArray tmpFields())
    
    Dim bytLoop As Byte
    
    For bytLoop = 0 To UBound(tmpFields)
        tmpFields(bytLoop).Editable = True
    Next bytLoop

End Function

Function ColorizeGrid(ParamArray tmpFields())
    
    Dim bytLoop As Byte
    
    For bytLoop = 0 To UBound(tmpFields)
        tmpFields(bytLoop).ForeColor = vbBlack
    Next bytLoop

End Function

Function EnableGrid(grid As iGrid, canEdit As Boolean, Optional myRow = 1, Optional myCol = 1)

    With grid
        .Enabled = True
        .Redraw = True
        .Editable = canEdit
        .RowMode = Not canEdit
        .SetCurCell myRow, myCol
    End With

End Function

Function EnableTabStop(ParamArray tmpFields())
    
    Dim bytLoop As Byte
    
    For bytLoop = 0 To UBound(tmpFields)
        tmpFields(bytLoop).TabStop = True
    Next bytLoop

End Function

Function LoadMessages()

    strMessages(1) = Chr(13) & "Το πεδίο είναι υποχρεωτικό." & Chr(13)
    strMessages(2) = Chr(13) & "Το πεδίο δεν είναι σωστό." & Chr(13)
    strMessages(3) = "Αν εγκαταλείψετε την επεξεργασία" & Chr(13) & "το αρχείο δεν θα ενημερωθεί." & Chr(13) & "Θέλετε σίγουρα να εγκαταλείψετε;"
    strMessages(4) = "Η εγγραφή θα διαγραφεί οριστικά." & Chr(13) & "Είστε σίγουροι ότι θέλετε" & Chr(13) & "να διαγράψετε την εγγραφή;"
    strMessages(5) = Chr(13) & "Δεν βρέθηκε εκτυπωτής αναφορών."
    strMessages(6) = Chr(13) & "Ο κωδικός υπάρχει ήδη." & Chr(13)
    strMessages(7) = Chr(13) & "Ο κωδικός δεν υπάρχει."
    strMessages(8) = Chr(13) & "Δεν βρέθηκαν εγγραφές."
    strMessages(9) = Chr(13) & "Δεν έχετε δώσει συναλλασόμενο."
    strMessages(10) = Chr(13) & "Η διαδικασία ολοκληρώθηκε."
    strMessages(11) = Chr(13) & "Η γραμμή "
    strMessages(12) = "Το πεδίο 'Νέος κωδικός' πρέπει" & Chr(13) & "να είναι ίδιο με" & Chr(13) & "το πεδίο 'Επιβεβαίωση κωδικού'."
    strMessages(13) = "ΣΥΝΕΧΕΙΑ ΑΠΟ ΠΡΟΗΓΟΥΜΕΝΗ ΣΕΛΙΔΑ"
    strMessages(14) = Chr(13) & "Η σχέση από - έως δεν είναι σωστή." & Chr(13)
    strMessages(15) = "Ο συνδυασμός " & Chr(13) & "χρήστης / κωδικός" & Chr(13) & "είναι λάθος."
    strMessages(16) = Chr(13) & "Η εφαρμογή εκτελείται ήδη."
    strMessages(17) = Chr(13) & "Η εγγραφή δεν βρέθηκε."
    strMessages(18) = "Ο εκτυπωτής που επιλέξατε δεν" & Chr(13) & "βρέθηκε στο σύστημα." & Chr(13) & "Ελέγξτε το όνομα και ξαναπροσπαθήστε."
    strMessages(19) = Chr(13) & "Θέλετε να τερματίσετε την εφαρμογή;" & Chr(13)
    strMessages(20) = "Διακοπή επεξεργασίας"
    strMessages(21) = "Νέα αναζήτηση"
    strMessages(22) = "Πρέπει να επανεκκινήσετε" & Chr(13) & "την εφαρμογή για να" & Chr(13) & "εφαρμοστούν οι αλλαγές που κάνατε."
    strMessages(23) = Chr(13) & "Η επεξεργασία διακόπηκε."
    strMessages(24) = "Η ΕΚΤΥΠΩΣΗ ΣΥΝΕΧΙΖΕΤΑΙ ΣΤΗΝ ΕΠΟΜΕΝΗ ΣΕΛΙΔΑ"
    strMessages(25) = "ΤΕΛΟΣ ΕΚΤΥΠΩΣΗΣ"
    strMessages(26) = "Η εργασία αντιμετώπισε πρόβλημα και δεν" & Chr(13) & " ολοκληρώθηκε. Ελέγξτε το αρχείο λαθών που έχει δημιουργηθεί."
    strMessages(27) = "Η εφαρμογή ξεκινάει. Εχετε λίγη υπομονή!"
    strMessages(28) = "Ο φορολογικός μηχανισμός" & Chr(13) & "δεν είναι ενεργός." & Chr(13) & "Θέλετε να συνεχίσετε;"
    strMessages(29) = Chr(13) & "Το παραστατικό που ακυρώνεται δεν βρέθηκε."
    strMessages(30) = "ΣΕ ΜΕΤΑΦΟΡΑ"
    strMessages(31) = "ΑΠΟ ΜΕΤΑΦΟΡΑ"
    strMessages(32) = "ΓΕΝΙΚΑ ΣΥΝΟΛΑ"
    strMessages(33) = Chr(13) & "Θέλετε να διακοπεί η επεξεργασία;"
    strMessages(34) = Chr(13) & "Ο έλεγχος ολοκληρώθηκε χωρίς να βρεθούν λάθη."
    strMessages(35) = "Ο έλεγχος ολοκληρώθηκε και" & Chr(13) & "βρέθηκαν λάθη." & Chr(13) & "Ελέγξτε το αρχείο λαθών που έχει δημιουργηθεί."
    strMessages(36) = "ΖΗΤΟΥΜΕΝΗ ΠΕΡΙΟΔΟΣ"
    strMessages(37) = Chr(13) & "Δεν έχετε δώσει είδος."
    strMessages(38) = "Ο έλεγχος ολοκληρώθηκε χωρίς λάθη," & Chr(13) & "αλλά προτείνονται αλλαγές." & Chr(13) & "Ελέγξτε το αρχείο που έχει δημιουργηθεί."
    strMessages(39) = "Πρέπει να ενημερώσετε το πεδίο" & Chr(13) & "'Αρθρο απαλλαγής ΦΠΑ'" & Chr(13) & "στα στοιχεία του πελάτη"
    
    strMessages(50) = "Η ημερομηνία έκδοσης αφορά" & Chr(13) & "διάστημα το οποίο είναι κλειδωμένο." & Chr(13) & "Θέλετε να συνεχίσετε;"
    strMessages(51) = Chr(13) & "Δεν έχετε επιλέξει παραστατικά."
    strMessages(52) = "Ο Α.Φ.Μ. που δώσατε" & Chr(13) & "δεν φαίνεται να είναι σωστός." & Chr(13) & "Θέλετε να συνεχίσετε;"
    strMessages(53) = "Ο Α.Φ.Μ. ανήκει στο συναλλασόμενο " & Chr(13)
    strMessages(54) = "Δεν έχετε στην αποθήκη αρκετό απόθεμα" & Chr(13) & "από το συγκεκριμένο είδος" & Chr(13) & "Θέλετε σίγουρα να συνεχίσετε;"
    strMessages(55) = Chr(13) & "Βρέθηκε λάθος στην αρίθμηση παραστατικών."
    strMessages(56) = "Δεν μπορείτε να εκδόσετε παραστατικό" & Chr(13) & "με ημερομηνία" & Chr(13) & "διαφορετική της σημερινής."
    strMessages(58) = "Επειδή το παραστατικό" & Chr(13) & "έχει εκτυπωθεί" & Chr(13) & "οι αλλαγές δεν επιτρέπονται."
    strMessages(59) = "Βεβαιωθείτε ότι για το παραστατικό δεν" & Chr(13) & "έχει γίνει κλείσιμο ημέρας ή ότι η ΕΑΦΔΣΣ" & Chr(13) & "βρίσκεται σε δοκιμαστική λειτουργία."
    strMessages(60) = "Δεν έχετε επιλέξει" & Chr(13) & "παραστατικά" & Chr(13) & "του ίδιου συναλλασόμενου."
    strMessages(61) = Chr(13) & "Δεν βρέθηκε εκτυπωτής παραστατικών."
    strMessages(62) = "Για γρηγορότερη αναζήτηση, πρέπει να" & Chr(13) & "δώσετε τουλάχιστον τρεις χαρακτήρες από" & Chr(13) & "την περιγραφή του είδους."
    strMessages(63) = Chr(13) & "Ο φορολογικός μηχανισμός δεν είναι ενεργός."
    strMessages(64) = Chr(13) & "Το παραστατικό είναι ήδη καταχωρημένο."
    
End Function

Function UpdateTableName(myRefersToID)

    Select Case myRefersToID
        Case Is = 0, 2
            UpdateTableName = "Suppliers"
        Case Is = 1, 3
            UpdateTableName = "Customers"
        Case Is = 4
            UpdateTableName = "Items"
    End Select

End Function

Function AddTotalsToOutputFile(myMessage, mySums, myTotals)

    Dim intLoop As Integer
    Dim IntegerFormat As String
    Dim FloatFormat As String
    
    Dim myColumns() As String
    
    myColumns = Split(myTotals, ",")
    
    Print #1, myMessage;
    
    'mySums = Array with amounts
    'myColumns = Array with columns. Example 999FY - F = Float mask, Y(es)/N(o) = Printable column
    For intLoop = 0 To UBound(myColumns)
        If Right(myColumns(intLoop), 1) = "Y" Then
            Print #1, _
                Tab(Left(myColumns(intLoop), 3) - Len(format(mySums(intLoop), FormatAsSelection(Mid(myColumns(intLoop), 4, 1))))); _
                format(mySums(intLoop), FormatAsSelection(Mid(myColumns(intLoop), 4, 1)));
        End If
    Next intLoop
    
    Print #1, ""

End Function

Function CheckForArrows(KeyCode)
    
    'Up
    If KeyCode = 38 Then
        Sendkeys "+{TAB}"
        KeyCode = 0
    End If
    
    'Down
    If KeyCode = 40 Then
        Sendkeys "{TAB}"
        KeyCode = 0
    End If
    
End Function

Function OpenDataBase(tmpCompany)

    On Error GoTo TrapError
    
    OpenDataBase = False
    
    strFullPathName = strPathName & tmpCompany
    Set CommonDB = DBEngine.OpenDataBase(strFullPathName, False, False)
    OpenDataBase = True
    Set dBaseTables = CommonDB.TableDefs
    
    Exit Function
    
TrapError:
    If Err.Number = 3031 Or Err.Number = 3029 Then
        Exit Function
    Else
        Exit Function
    End If
    
End Function

Function PrinterExists(strPrinterName)

    Dim strPrinter As Printer
    
    For Each strPrinter In Printers
        If LCase(strPrinter.DeviceName) = LCase(strPrinterName) Then
            Set Printer = strPrinter
            PrinterExists = True
            Exit For
        End If
    Next

End Function

Function ResetKeyCode(KeyCode As Integer, Shift As Integer)

    Dim CtrlDown
    
    CtrlDown = Shift + vbCtrlMask
    
    If _
        (KeyCode = vbKeyEscape) Or _
        (KeyCode = vbKeyF4) Or _
        (KeyCode = vbKeyN And CtrlDown > 2) Or _
        (KeyCode = vbKeyE And CtrlDown > 2) Or _
        (KeyCode = vbKeyS And CtrlDown > 2) Or _
        (KeyCode = vbKeyD And CtrlDown > 2) Or _
        (KeyCode = vbKeyP And CtrlDown > 2) Or _
        (KeyCode = vbKeyC And CtrlDown > 2) Or _
        (KeyCode = vbKeyL And CtrlDown > 2) Or _
        (KeyCode = vbKeyF And CtrlDown > 2) Then KeyCode = 0
    
    ResetKeyCode = KeyCode
    
End Function

Function AddDummyLines(grdGrid, ParamArray columns() As Variant)

    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngLoop As Long
    
    grdGrid.Redraw = False
    
    For lngRow = 1 To 40
        With grdGrid
            .AddRow
            For lngCol = 0 To (UBound(columns))
                .CellValue(lngRow, lngCol + 1) = String(columns(lngCol), "A")
            Next lngCol
        End With
    Next lngRow
    
    grdGrid.Redraw = True

End Function

Function AddColumnsToGrid(grdGrid As iGrid, headerHeight, strLayoutCol, tmpElements, tmpTitles)

    On Error GoTo ErrTrap
    
    Dim intLoop As Integer
    Dim intNoOfElements As Integer
    Dim strKey As String
    Dim strHeader As String
    Dim intOuter As Integer
    Dim lngCol As Long
    
    intNoOfElements = 0
    
    With grdGrid
        .Clear True
        .Redraw = False
        .GridLines = igGridLinesNone
        .Visible = False
    End With
    
    ReDim arrWidth(1)
    ReDim arrJustification(1)
    ReDim arrFormat(1)
    ReDim arrKey(1)
    ReDim arrAllowSizing(1)
    ReDim arrHeaderTitle(1)
    
    For intOuter = 1 To Len(tmpElements)
        intNoOfElements = intNoOfElements + 1
        'Πλάτος
        ReDim Preserve arrWidth(intNoOfElements)
        arrWidth(intNoOfElements) = Mid(tmpElements, intOuter, 2)
        intOuter = intOuter + 2
        'Επιτρέπεται η αλλαγή πλάτους
        ReDim Preserve arrAllowSizing(intNoOfElements)
        arrAllowSizing(intNoOfElements) = Mid(tmpElements, intOuter, 1)
        intOuter = intOuter + 1
        'Στοίχιση
        ReDim Preserve arrJustification(intNoOfElements)
        arrJustification(intNoOfElements) = Mid(tmpElements, intOuter, 1)
        intOuter = intOuter + 1
        'Μορφή
        ReDim Preserve arrFormat(intNoOfElements)
        arrFormat(intNoOfElements) = Mid(tmpElements, intOuter, 1)
        intOuter = intOuter + 1
        'ColKey
        ReDim Preserve arrKey(intNoOfElements)
        Do Until Mid(tmpElements, intOuter, 1) = ","
            If intOuter <= Len(tmpElements) Then
                strKey = strKey + Mid(tmpElements, intOuter, 1)
                intOuter = intOuter + 1
            Else
                Exit Do
            End If
        Loop
        arrKey(intNoOfElements) = strKey
        strKey = ""
    Next intOuter
    
    intNoOfElements = 0
    
    For intOuter = 1 To Len(tmpTitles)
        intNoOfElements = intNoOfElements + 1
        ReDim Preserve arrHeaderTitle(intNoOfElements)
        Do Until Mid(tmpTitles, intOuter, 1) = ","
            If intOuter <= Len(tmpTitles) Then
                strHeader = strHeader + Mid(tmpTitles, intOuter, 1)
                intOuter = intOuter + 1
            Else
                Exit Do
            End If
        Loop
        arrHeaderTitle(intNoOfElements) = strHeader
        strHeader = ""
    Next intOuter

    For intLoop = 1 To intNoOfElements
        strHeader = arrHeaderTitle(intLoop)
        With grdGrid.AddCol(sKey:=IIf(Left(arrKey(intLoop), 1) <> "X", arrKey(intLoop), Right(arrKey(intLoop), Len(arrKey(intLoop)) - 1)), sHeader:=strHeader, lWidth:=arrWidth(intLoop), eHdrTextFlags:=igTextCenter)
            Select Case arrJustification(intLoop)
                Case "L": .eTextFlags = 0
                Case "C": .eTextFlags = 1
                Case "R": .eTextFlags = 2
            End Select
            Select Case arrFormat(intLoop)
                Case "I"
                    .sFmtString = "#,##0"
                Case "F"
                    .sFmtString = "#,##0.00"
                Case "D"
                    .sFmtString = "dd/mm/yyyy"
                Case "T"
                    .sFmtString = "hh:mm"
            End Select
        End With
        grdGrid.ColHeaderTextFlags(intLoop) = 32821
        grdGrid.ColTag(intLoop) = arrAllowSizing(intLoop)
        If Left(arrKey(intLoop), 1) = "X" Then
            grdGrid.ColHeaderTextFlags(intLoop) = 32789
        End If
    Next intLoop
    
    With grdGrid
        .LayoutCol = strLayoutCol
        .Header.Height = headerHeight
        .Redraw = True
        .Visible = True
    End With
    
    Exit Function

ErrTrap:
    AddColumnsToGrid = False
    DisplayErrorMessage True, Err.Description
    
    Exit Function

End Function

Function PositionControls(thisForm As Form, formFullScreen As Boolean, Optional grdGrid As iGrid)

    On Error GoTo ErrTrap
    
    Dim ctl As Control
    Dim intLoop As Integer
    
    intLoop = 0
    
    'Ενα - ένα
    For Each ctl In thisForm.Controls
        'Τα κάνει αόρατα
        If ctl.Name = "frmInfo" Then
            thisForm.frmInfo.Visible = False
        End If
        'Κουμπιά
        If ctl.Name = "cmdButton" Then
            thisForm.cmdButton(intLoop).ButtonStyle = gbOfficeXP
            intLoop = intLoop + 1
        End If
    Next ctl
    
    'Πλήρης οθόνη
    If formFullScreen Then PositionFullScreenControls thisForm, True, grdGrid
    
    'Οχι πλήρης οθόνη
    If Not formFullScreen Then PositionCenteredScreenControls thisForm, True, grdGrid
    
    'Πρόοδος
    For Each ctl In thisForm.Controls
        If ctl.Name = "frmProgress" Then
            With thisForm.frmProgress
                .Visible = False
                .ZOrder 1
                .Top = (thisForm.Height / 2) - (.Height / 2)
                .Left = (thisForm.Width / 2) - (.Width / 2)
                Exit For
            End With
        End If
        If ctl.Name = "frmTotals" Then
            With thisForm.frmTotals
                .Left = (thisForm.frmContainer.Width / 2) - (.Width / 2)
            End With
        End If
    Next ctl
    
    'Σημερινή ημερομηνία
    For Each ctl In thisForm.Controls
        If ctl.Name = "lblToday" Then thisForm.lblToday.Caption = format(Date, "dddd dd/mm/yyyy")
    Next ctl

    Exit Function
    
ErrTrap:
    If Err.Number = 438 Then Resume Next 'Το αντικείμενο δεν υπάρχει

End Function

Function CountRecords(myRecordset As Recordset)

    CountRecords = 0
    
    If Not myRecordset.EOF Then
        myRecordset.MoveLast
        CountRecords = myRecordset.RecordCount
        myRecordset.MoveFirst
        Exit Function
    End If

End Function

Function ToggleInfoPanel(thisForm As Form)

    With thisForm.frmInfo
        If .Visible = True Then
            .Visible = False
        Else
            .Visible = True
            .Left = 2500
            .Top = 3000
            .ZOrder 0
        End If
    End With

End Function

Function CheckForLoadedForm(thisForm As String)

    Dim intIndex As Integer
    Dim f As Form
    Dim myForms
    
    intIndex = 0
    myForms = Split(thisForm, ",")
    
    On Error Resume Next
    
    For Each f In Forms
        For intIndex = 0 To UBound(myForms)
            If f.Name = myForms(intIndex) Then
                CheckForLoadedForm = True
                Exit For
            End If
        Next intIndex
    Next f
    
End Function
