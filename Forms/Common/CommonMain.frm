VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{77EBD0B1-871A-4AD1-951A-26AEFE783111}#2.1#0"; "vbalExpBar6.ocx"
Begin VB.Form CommonMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   Caption         =   " Εμπορική Διαχείρηση"
   ClientHeight    =   5790
   ClientLeft      =   165
   ClientTop       =   210
   ClientWidth     =   7140
   Icon            =   "CommonMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   7140
   WindowState     =   2  'Maximized
   Begin vbalExplorerBarLib6.vbalExplorerBarCtl vbExplorerBar 
      Height          =   540
      Left            =   6525
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   4500
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   953
      BackColorEnd    =   32896
      BackColorStart  =   16744576
   End
   Begin vbalIml6.vbalImageList ilsExplorerIcons 
      Left            =   6525
      Top             =   5100
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   32
      IconSizeY       =   32
      ColourDepth     =   8
      Size            =   35296
      Images          =   "CommonMain.frx":1CCA
      Version         =   131072
      KeyCount        =   8
      Keys            =   "ÿÿÿÿÿÿÿ"
   End
   Begin VB.Image imgImage 
      Appearance      =   0  'Flat
      Height          =   5910
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7245
   End
End
Attribute VB_Name = "CommonMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub ResizeBar(intKey, blnState As Boolean, ExplorerBar As vbalExplorerBarCtl, ParamArray Buttons() As Variant)

    With ExplorerBar
        .Height = GetSetting(strAppTitle, "Appearance", "Menu Height")
        If Not blnState Then .Top = (Me.Height / 2) - (.Height / 2): Exit Sub
        .Redraw = False
        .Height = Buttons(intKey - 1)
        .Top = (Me.Height / 2) - (.Height / 2) - 50
        .Redraw = True
    End With

End Sub

Sub BuildMenu()

    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem
    
    Dim intLoop As Integer
    Dim intItem As Integer
    Dim strMenuCategory As String
    Dim strMenuCategories As String
    
    With CommonMain
        .Tag = "True"
        .Height = Screen.Height
        .ScaleHeight = .Height
    End With
    
    strMenuCategories = GetSetting(strAppTitle, "Appearance", "Menu Categories")
    For intLoop = 1 To Len(strMenuCategories)
        While Mid(strMenuCategories, intLoop, 1) <> "," And intLoop <= Len(strMenuCategories)
            strMenuCategory = strMenuCategory & Mid(strMenuCategories, intLoop, 1)
            intLoop = intLoop + 1
        Wend
        intItem = intItem + 1
        ReDim Preserve arrMenu(intItem)
        arrMenu(intItem) = Int(strMenuCategory)
        strMenuCategory = ""
    Next intLoop
    
    With CommonMain.vbExplorerBar
        .BarTitleImageList = CommonMain.ilsExplorerIcons.hIml
        .Height = GetSetting(strAppTitle, "Appearance", "Menu Height")
        .Left = ((Screen.Width / Screen.TwipsPerPixelX) / 3)
        .Redraw = False
        .Top = (CommonMain.Height / 2) - (.Height / 2) - 200
        .UseExplorerStyle = False
        .Width = GetSetting(strAppTitle, "Appearance", "Menu Width")
        .BackColorStart = vbBlack
        .BackColorEnd = vbBlack
        'Αγορές
        Set cBar = .Bars.Add(, "Αγορές", "  Αγορές")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        cBar.IconIndex = 0
        Set cItem = cBar.Items.Add(, "ΑγορέςΚινήσεις", " - Κινήσεις")
        Set cItem = cBar.Items.Add(, "ΑγορέςΗμερολόγιοΚινήσεων", " - Ημερολόγιο κινήσεων")
        Set cItem = cBar.Items.Add(, "ΑγορέςΕκκρεμήΔελτίαΑποστολής", " - Εκκρεμή δελτία αποστολής")
        Set cItem = cBar.Items.Add(, "ΑγορέςΤύποιΠαραστατικών", " - Τύποι παραστατικών")
        'Πωλήσεις
        Set cBar = .Bars.Add(, "Πωλήσεις", "  Πωλήσεις")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        cBar.IconIndex = 1
        Set cItem = cBar.Items.Add(, "ΠωλήσειςΚινήσεις", " - Κινήσεις")
        Set cItem = cBar.Items.Add(, "SalesIncomingVehicles", " - Ημερολόγιο εισερχομένων οχημάτων")
        Set cItem = cBar.Items.Add(, "ΠωλήσειςΗμερολόγιοΚινήσεων", " - Ημερολόγιο κινήσεων")
        Set cItem = cBar.Items.Add(, "ΠωλήσειςΕκκρεμήΔελτίαΑποστολής", " - Εκκρεμή δελτία αποστολής")
        Set cItem = cBar.Items.Add(, "ΠωλήσειςΤιμολόγησηΤριγωνικών", " - Τιμολόγηση τριγωνικών πωλήσεων")
        Set cItem = cBar.Items.Add(, "ΠωλήσειςΤύποιΠαραστατικών", " - Τύποι παραστατικών")
        'Είδη
        Set cBar = .Bars.Add(, "Είδη", "  Είδη")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        cBar.IconIndex = 2
        Set cItem = cBar.Items.Add(, "Items", " - Διαχείρηση")
        Set cItem = cBar.Items.Add(, "ItemsIndex", " - Ευρετήριο")
        Set cItem = cBar.Items.Add(, "ItemsTransactions", " - Κινήσεις")
        Set cItem = cBar.Items.Add(, "ΕίδηΗμερολόγιοΚινήσεων", " - Ημερολόγιο κινήσεων")
        Set cItem = cBar.Items.Add(, "ItemsLedger", " - Καρτέλα")
        Set cItem = cBar.Items.Add(, "itemsBalanceSheet", " - Ισοζύγιο")
        Set cItem = cBar.Items.Add(, "ItemsInventory", " - Υπόλοιπα")
        Set cItem = cBar.Items.Add(, "ΕίδηΤύποιΠαραστατικών", " - Τύποι παραστατικών")
        'Πελάτες
        Set cBar = .Bars.Add(, "Πελάτες", "  Πελάτες")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        cBar.IconIndex = 3
        Set cItem = cBar.Items.Add(, "ΠελάτεςΔιαχείρηση", " - Διαχείρηση")
        Set cItem = cBar.Items.Add(, "ΠελάτεςΕυρετήριο", " - Ευρετήριο")
        Set cItem = cBar.Items.Add(, "ΠελάτεςΚινήσεις", " - Κινήσεις")
        Set cItem = cBar.Items.Add(, "ΠελάτεςΗμερολόγιοΚινήσεων", " - Ημερολόγιο κινήσεων")
        Set cItem = cBar.Items.Add(, "ΠελάτεςΚαρτέλα", " - Καρτέλα")
        Set cItem = cBar.Items.Add(, "ΠελάτεςΙσοζύγιο", " - Ισοζύγιο")
        Set cItem = cBar.Items.Add(, "ΠελάτεςΕισπρακτέαΑξιόγραφα", " - Ημερολόγιο εισπρακτέων αξιογράφων")
        Set cItem = cBar.Items.Add(, "ΠελάτεςΤύποιΠαραστατικών", " - Τύποι παραστατικών")
        'Προμηθευτές
        Set cBar = .Bars.Add(, "Προμηθευτές", "  Προμηθευτές")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        cBar.IconIndex = 4
        Set cItem = cBar.Items.Add(, "ΠρομηθευτέςΔιαχείρηση", " - Διαχείρηση")
        Set cItem = cBar.Items.Add(, "ΠρομηθευτέςΕυρετήριο", " - Ευρετήριο")
        Set cItem = cBar.Items.Add(, "ΠρομηθευτέςΚινήσεις", " - Κινήσεις")
        Set cItem = cBar.Items.Add(, "ΠρομηθευτέςΗμερολόγιοΚινήσεων", " - Ημερολόγιο κινήσεων")
        Set cItem = cBar.Items.Add(, "ΠρομηθευτέςΚαρτέλα", " - Καρτέλα")
        Set cItem = cBar.Items.Add(, "ΠρομηθευτέςΙσοζύγιο", " - Ισοζύγιο")
        Set cItem = cBar.Items.Add(, "ΠρομηθευτέςΠληρωτέαΑξιόγραφα", " - Ημερολόγιο πληρωτέων αξιογράφων")
        Set cItem = cBar.Items.Add(, "ΠρομηθευτέςΤύποιΠαραστατικών", " - Τύποι παραστατικών")
        'ΦΠΑ
        Set cBar = .Bars.Add(, "ΦΠΑ", "  Φ.Π.Α.")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        cBar.IconIndex = 5
        Set cItem = cBar.Items.Add(, "VATBalanceSheet", " - Ισοζύγιο Φ.Π.Α.")
        'Βοηθητικά
        Set cBar = .Bars.Add(, "Βοηθητικά", "  Βοηθητικά")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        cBar.IconIndex = 6
        Set cItem = cBar.Items.Add(, "Παραμετροποίηση", "Παραμετροποίηση")
            cItem.ItemType = eItemText
            cItem.Bold = True
            cItem.TextColor = RGB(96, 150, 207)
            Set cItem = cBar.Items.Add(, "UtilsSettings", Space(5) & " - Γενικές παράμετροι")
            Set cItem = cBar.Items.Add(, "UtilsPrinters", Space(5) & " - Εκτυπωτές")
        Set cItem = cBar.Items.Add(, "Πίνακες", "Πίνακες")
            cItem.ItemType = eItemText
            cItem.Bold = True
            cItem.TextColor = RGB(96, 150, 207)
            Set cItem = cBar.Items.Add(, "UtilsManufacturers", Space(5) & " - Κατασκευαστές")
            Set cItem = cBar.Items.Add(, "UtilsItemCategories", Space(5) & " - Κατηγορίες ειδών")
            Set cItem = cBar.Items.Add(, "UtilsTaxOffices", Space(5) & " - Οικονομικές υπηρεσίες")
            Set cItem = cBar.Items.Add(, "UtilsDeliveryPoints", Space(5) & " - Τόποι παραλαβής")
            Set cItem = cBar.Items.Add(, "UtilsBanks", Space(5) & " - Τράπεζες")
            Set cItem = cBar.Items.Add(, "UtilsPaymentWays", Space(5) & " - Τρόποι πληρωμής")
            Set cItem = cBar.Items.Add(, "UtilsUsers", Space(5) & " - Χρήστες")
            Set cItem = cBar.Items.Add(, "UtilsCountries", Space(5) & " - Χώρες")
        Set cItem = cBar.Items.Add(, "Εργασίες", "Εργασίες")
            cItem.ItemType = eItemText
            cItem.Bold = True
            cItem.TextColor = RGB(96, 150, 207)
            Set cItem = cBar.Items.Add(, "UtilsTablesCheck", Space(5) & " - Ελεγχος αρχείων")
            Set cItem = cBar.Items.Add(, "UtilsUpdateItemQty", Space(5) & " - Ενημέρωση ποσοτήτων")
        'Τερματισμός
        Set cBar = .Bars.Add(, "Εξοδος", "  Εξοδος")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        cBar.IconIndex = 7
        Set cItem = cBar.Items.Add(, "ΕξοδοςΑλλαγήΕταιρίας", "- Αλλαγή εταιρίας")
        Set cItem = cBar.Items.Add(, "ΕξοδοςΤερματισμόςΕφαρμογής", "- Τερματισμός εφαρμογής")
        .Redraw = True
    End With

End Sub

Function CloseApp()

    If MyMsgBox(2, strAppTitle, strMessages(19), 2) Then CloseApp = True

End Function

Private Sub Form_Load()

    On Error GoTo ErrTrap
    
    strImageDirectory = GetSetting(strAppTitle, "Path Names", "Image Directory")
    
    If strImageDirectory <> "" Then imgImage.Picture = LoadPicture(strImageDirectory & "Background.jpg")
    If strImageDirectory <> "" Then CommonMain.Icon = LoadPicture(strImageDirectory & "Icon.ico")
    
    BuildMenu
    
    With CommonMain
        .ScaleHeight = .Height
        .ScaleWidth = .Width
        .imgImage.Height = Screen.Height - 1000
        .imgImage.Top = (.Height / 2) - (.imgImage.Height / 2) - 200
        .imgImage.Left = .vbExplorerBar.Left * 2 + .vbExplorerBar.Width - 200
        .imgImage.Width = Screen.Width - (.vbExplorerBar.Left * 3) - .vbExplorerBar.Width + 400
        .BackColor = vbBlack
        .Refresh
    End With
    
    ResizeBar 1, 1, vbExplorerBar, arrMenu(1), arrMenu(2), arrMenu(3), arrMenu(4), arrMenu(5), arrMenu(6), arrMenu(7), arrMenu(8)
    
    strUnicodeFile = strReportsPathName & GetSetting(strAppTitle, "Path Names", "UnicodeFileName")
    strAsciiFile = GetSetting(strAppTitle, "Path Names", "AsciiFileName")
    
    Exit Sub
    
ErrTrap:
    If Err.Number = 380 Then Exit Sub
    If Err.Number = 53 Then Resume Next

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim obj As Object
    
    'Επιλογή κλεισίματος απο το μενού συστήματος, κλικ στο Χ ή ALT-F4
    If UnloadMode = 0 Then
        If CloseApp Then
            For Each obj In Forms
                Unload obj
            Next
            KillProcess strApplicationEXEName: End
        Else
            Cancel = 1
            Exit Sub
        End If
    End If
    
    'Επιλογή κλεισίματος από την επιλογή Εξοδος > Τερματισμός
    If UnloadMode = 1 Then
        KillProcess strApplicationEXEName
    End If

End Sub

Private Sub vbExplorerBar_BarClick(bar As vbalExplorerBarLib6.cExplorerBar)

    ResizeBar bar.Index, bar.State, vbExplorerBar, arrMenu(1), arrMenu(2), arrMenu(3), arrMenu(4), arrMenu(5), arrMenu(6), arrMenu(7), arrMenu(8)

End Sub

Private Sub vbExplorerBar_ItemClick(itm As vbalExplorerBarLib6.cExplorerBarItem)

    Dim obj As Object

    Select Case itm.Key
        'Αγορές
        Case "ΑγορέςΚινήσεις"
            With CommonTransactions
                .txtRefersTo.text = "1"
                .lblTitle.Caption = "Αγορές"
                .txtTable.text = "Suppliers"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΑγορέςΗμερολόγιοΚινήσεων"
            With CommonTransactionsIndex
                .lblTitle.Caption = "Ημερολόγιο αγορών"
                .txtTable.text = "Suppliers"
                .txtOppositeTable.text = ""
                .txtRefersTo.text = "1"
                .txtOppositeRefersTo.text = ""
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΑγορέςΕκκρεμήΔελτίαΑποστολής"
            With CommonPendingInvoices
                .lblTitle.Caption = "Εκκρεμή δελτία αποστολής αγορών"
                .txtTable.text = "Suppliers"
                .txtRefersTo.text = "1"
                .txtInitialRefersTo.text = "1"
                .txtTriangularID.text = "0"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΑγορέςΤύποιΠαραστατικών"
            With UtilsCodes
                .lblTitle.Caption = "Τύποι παραστατικών αγορών"
                .txtRefersTo.text = "1"
                .Tag = "True"
                .Show 1, Me
            End With
            
        'Πωλήσεις
        Case "ΠωλήσειςΚινήσεις"
            With CommonTransactions
                .txtRefersTo.text = "2"
                .lblTitle.Caption = "Πωλήσεις"
                .txtTable.text = "Customers"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠωλήσειςΗμερολόγιοΚινήσεων"
            With CommonTransactionsIndex
                .lblTitle.Caption = "Ημερολόγιο πωλήσεων"
                .txtTable.text = "Customers"
                .txtOppositeTable.text = ""
                .txtRefersTo.text = "2"
                .txtOppositeRefersTo.text = ""
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠωλήσειςΕκκρεμήΔελτίαΑποστολής"
            With CommonPendingInvoices
                .lblTitle.Caption = "Εκκρεμή δελτία αποστολής πωλήσεων"
                .txtTable.text = "Customers"
                .txtRefersTo.text = "2"
                .txtInitialRefersTo.text = "2"
                .txtTriangularID.text = "0"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "SalesIncomingVehicles"
            With SalesIncomingVehicles
                .lblTitle.Caption = "Ημερολόγιο εισερχομένων οχημάτων"
                .txtTable.text = "Customers"
                .txtRefersTo.text = "2"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠωλήσειςΤιμολόγησηΤριγωνικών"
            With CommonPendingInvoices
                .lblTitle.Caption = "Τιμολόγηση τριγωνικών πωλήσεων"
                .txtTable.text = "Suppliers"
                .txtRefersTo.text = "1"
                .txtInitialRefersTo.text = "1"
                .txtTriangularID.text = "1"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠωλήσειςΤύποιΠαραστατικών"
            With UtilsCodes
                .lblTitle.Caption = "Τύποι παραστατικών πωλήσεων"
                .txtRefersTo.text = "2"
                .Tag = "True"
                .Show 1, Me
            End With
        
        'Είδη
        Case "Items"
            With Items
                .txtTable.text = "Items"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ItemsIndex"
            With ItemsIndex
                .txtTable.text = "Items"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ItemsTransactions"
            With ItemsTransactions
                .txtRefersTo.text = "5"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΕίδηΗμερολόγιοΚινήσεων"
            With CommonTransactionsIndex
                .lblTitle.Caption = "Ημερολόγιο ειδών"
                .txtTable.text = "Items"
                .txtOppositeTable.text = ""
                .txtRefersTo.text = "5"
                .txtOppositeRefersTo.text = ""
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ItemsLedger"
            With ItemsLedger
                .txtTable.text = "Items"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "itemsBalanceSheet"
            With itemsBalanceSheet
                .txtTable.text = "Items"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ItemsInventory"
            With ItemsInventory
                .txtTable.text = "Items"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΕίδηΤύποιΠαραστατικών"
            With UtilsCodes
                .lblTitle.Caption = "Τύποι παραστατικών ειδών"
                .txtRefersTo.text = "5"
                .Tag = "True"
                .Show 1, Me
            End With
        
        'Πελάτες
        Case "ΠελάτεςΔιαχείρηση"
            With Persons 'OK
                .txtTable.text = "Customers"
                .txtOppositeTable.text = "Suppliers"
                .txtRefersTo.text = "4"
                .lblTitle.Caption = "Πελάτες"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠελάτεςΕυρετήριο"
            With PersonsIndex 'OK
                .txtTable.text = "Customers"
                .txtOppositeTable.text = "Suppliers"
                .txtRefersTo.text = "4"
                .lblTitle.Caption = "Ευρετήριο πελατών"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠελάτεςΚινήσεις"
            With PersonsTransactions 'OK
                .txtTable.text = "Customers"
                .txtOppositeTable.text = "Suppliers"
                .txtRefersTo.text = "4"
                .lblTitle.Caption = "Κινήσεις πελατών"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠελάτεςΗμερολόγιοΚινήσεων"
            With CommonTransactionsIndex 'OK
                .lblTitle.Caption = "Ημερολόγιο κινήσεων πελατών"
                .txtTable.text = "Customers"
                .txtRefersTo.text = "4"
                .txtOppositeTable.text = "Suppliers"
                .txtOppositeRefersTo.text = "2"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠελάτεςΚαρτέλα"
            With PersonsLedger 'OK
                .txtRefersTo.text = "4"
                .lblTitle.Caption = "Καρτέλα πελάτη"
                .txtTable.text = "Customers"
                .txtOppositeTable.text = "Suppliers"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠελάτεςΙσοζύγιο"
            With PersonsBalanceSheet 'OK
                .lblTitle.Caption = "Ισοζύγιο πελατών"
                .txtTable.text = "Customers"
                .txtOppositeTable.text = "Suppliers"
                .txtRefersTo.text = "4"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠελάτεςΕισπρακτέαΑξιόγραφα"
            With PersonsChecksIndex 'OK
                .lblTitle.Caption = "Ημερολόγιο εισπρακτέων αξιογράφων"
                .txtTable.text = "Customers"
                .txtRefersTo.text = "4"
                .txtOppositeTable.text = "Suppliers"
                .txtOppositeRefersTo.text = "3"
                .txtIssuedBy.text = "Εκδότης"
                .txtHoldedBy.text = "Κάτοχος"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠελάτεςΤύποιΠαραστατικών"
            With UtilsCodes 'OK
                .lblTitle.Caption = "Τύποι παραστατικών πελατών"
                .txtRefersTo.text = "4"
                .Tag = "True"
                .Show 1, Me
            End With
        
        'Προμηθευτές
        Case "ΠρομηθευτέςΔιαχείρηση"
            With Persons 'OK
                .txtTable.text = "Suppliers"
                .txtOppositeTable.text = "Customers"
                .txtRefersTo.text = "3"
                .lblTitle.Caption = "Προμηθευτές"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠρομηθευτέςΕυρετήριο"
            With PersonsIndex 'OK
                .txtTable.text = "Suppliers"
                .txtOppositeTable.text = "Customers"
                .txtRefersTo.text = "3"
                .lblTitle.Caption = "Ευρετήριο προμηθευτών"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠρομηθευτέςΚινήσεις"
            With PersonsTransactions 'OK
                .txtTable.text = "Suppliers"
                .txtOppositeTable.text = "Customers"
                .txtRefersTo.text = "3"
                .lblTitle.Caption = "Κινήσεις προμηθευτών"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠρομηθευτέςΗμερολόγιοΚινήσεων"
            With CommonTransactionsIndex 'OK
                .lblTitle.Caption = "Ημερολόγιο κινήσεων προμηθευτών"
                .txtTable.text = "Suppliers"
                .txtRefersTo.text = "3"
                .txtOppositeTable.text = "Customers"
                .txtOppositeRefersTo.text = "4"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠρομηθευτέςΚαρτέλα"
            With PersonsLedger 'OK
                .txtRefersTo.text = "3"
                .lblTitle.Caption = "Καρτέλα προμηθευτή"
                .txtTable.text = "Suppliers"
                .txtOppositeTable.text = "Customers"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠρομηθευτέςΙσοζύγιο"
            With PersonsBalanceSheet 'OK
                .lblTitle.Caption = "Ισοζύγιο προμηθευτών"
                .txtTable.text = "Suppliers"
                .txtOppositeTable.text = "Customers"
                .txtRefersTo.text = "3"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠρομηθευτέςΠληρωτέαΑξιόγραφα"
            With PersonsChecksIndex 'OK
                .lblTitle.Caption = "Ημερολόγιο πληρωτέων αξιογράφων"
                .txtTable.text = "Suppliers"
                .txtRefersTo.text = "3"
                .txtOppositeTable.text = "Customers"
                .txtOppositeRefersTo.text = "4"
                .txtIssuedBy.text = "Κάτοχος"
                .txtHoldedBy.text = "Εκδότης"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠρομηθευτέςΤύποιΠαραστατικών"
            With UtilsCodes 'OK
                .lblTitle.Caption = "Τύποι παραστατικών προμηθευτών"
                .txtRefersTo.text = "3"
                .Tag = "True"
                .Show 1, Me
            End With
            
        'ΦΠΑ
        Case "VATBalanceSheet"
            With VATBalanceSheet
                .Tag = "True"
                .Show 1, Me
            End With
        
        'Βοηθητικά
        Case "UtilsSettings"
            With UtilsSettings
                .Tag = "True"
                .Show 1, Me
            End With
        Case "UtilsPrinters"
            With UtilsPrinters
                .Tag = "True"
                .Show 1, Me
            End With
        
        Case "UtilsManufacturers"
            With UtilsManufacturers
                .Tag = "True"
                .Show 1, Me
            End With
        Case "UtilsItemCategories"
            With UtilsItemCategories
                .Tag = "True"
                .Show 1, Me
            End With
        Case "UtilsTaxOffices"
            With UtilsTaxOffices
                .Tag = "True"
                .Show 1, Me
            End With
        Case "UtilsDeliveryPoints"
            With UtilsDeliveryPoints
                .Tag = "True"
                .Show 1, Me
            End With
        Case "UtilsPaymentWays"
            With UtilsPaymentWays
                .Tag = "True"
                .Show 1, Me
            End With
        Case "UtilsBanks"
            With UtilsBanks
                .Tag = "True"
                .Show 1, Me
            End With
        Case "UtilsUsers"
            With UtilsUsers
                .Tag = "True"
                .Show 1, Me
            End With
        Case "UtilsCountries"
            With UtilsCountries
                .Tag = "True"
                .Show 1, Me
            End With
            
        Case "UtilsTablesCheck"
            With UtilsTablesCheck
                .Tag = "True"
                .Show 1, Me
            End With
            
        Case "UtilsUpdateItemQty"
            With UtilsUpdateItemQty
                .Tag = "True"
                .Show 1, Me
            End With
            
        'Εξοδος
        Case "ΕξοδοςΑλλαγήΕταιρίας"
            CommonLogin.Tag = "True"
            CommonLogin.Show
        Case "ΕξοδοςΤερματισμόςΕφαρμογής"
            If CloseApp Then
                For Each obj In Forms
                    Unload obj
                Next
            End If
        
    End Select

End Sub

