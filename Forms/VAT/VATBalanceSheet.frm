VERSION 5.00
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form VATBalanceSheet 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   9765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19125
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   19125
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   375
      TabIndex        =   10
      Top             =   5550
      Width           =   4590
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         Caption         =   "Συνέχεια"
         ButtonStyle     =   2
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   99
         ShowFocusRect   =   0   'False
         BackColor       =   8438015
      End
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   2
         Left            =   3075
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         Caption         =   "Κλείσιμο"
         ButtonStyle     =   2
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   99
         ShowFocusRect   =   0   'False
         BackColor       =   8421631
      End
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   1
         Left            =   1650
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         Caption         =   "Νέα αναζήτηση"
         ButtonStyle     =   2
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   99
         ShowFocusRect   =   0   'False
         BackColor       =   8438015
      End
   End
   Begin VB.Frame frmCriteria 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3915
      Index           =   0
      Left            =   450
      TabIndex        =   5
      Top             =   1125
      Width           =   5190
      Begin UserControls.newDate mskIssueFrom 
         Height          =   465
         Left            =   1650
         TabIndex        =   0
         Top             =   825
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   820
         ForeColor       =   0
         Text            =   ""
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newDate mskIssueTo 
         Height          =   465
         Left            =   3225
         TabIndex        =   1
         Top             =   825
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   820
         ForeColor       =   0
         Text            =   ""
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newFloat mskNetBuys 
         Height          =   465
         Left            =   1650
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   820
         Enabled         =   0   'False
         Alignment       =   1
         ForeColor       =   0
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newFloat mskVATBuys 
         Height          =   465
         Left            =   3225
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   820
         Enabled         =   0   'False
         Alignment       =   1
         ForeColor       =   0
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newFloat mskVATSales 
         Height          =   465
         Left            =   3225
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2025
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   820
         Enabled         =   0   'False
         Alignment       =   1
         ForeColor       =   0
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newFloat mskNetSales 
         Height          =   465
         Left            =   1650
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2025
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   820
         Enabled         =   0   'False
         Alignment       =   1
         ForeColor       =   0
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newFloat mskVATBalance 
         Height          =   465
         Left            =   3225
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2700
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   820
         Enabled         =   0   'False
         Alignment       =   1
         ForeColor       =   0
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Υπόλοιπο"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   0
         Left            =   450
         TabIndex        =   22
         Top             =   2775
         Width           =   2190
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000000C0&
         Caption         =   "Πωλήσεις"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   6
         Left            =   450
         TabIndex        =   21
         Top             =   2100
         Width           =   765
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000000C0&
         Caption         =   "Αγορές"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   450
         TabIndex        =   20
         Top             =   1575
         Width           =   765
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000000C0&
         Caption         =   "Εκδοση"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   450
         TabIndex        =   14
         Top             =   900
         Width           =   765
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         Caption         =   "Κριτήρια αναζήτησης"
         BeginProperty Font 
            Name            =   "Aka-Acid-Steelfish"
            Size            =   14.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Index           =   3
         Left            =   150
         TabIndex        =   8
         Top             =   75
         Width           =   1665
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   2
         Left            =   0
         Top             =   750
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   0
         Left            =   4725
         Top             =   1125
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   1
         Left            =   1200
         Top             =   1275
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblToday 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808000&
         Caption         =   "01/05/2017"
         BeginProperty Font 
            Name            =   "Aka-Acid-Steelfish"
            Size            =   14.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   525
         TabIndex        =   6
         Top             =   75
         Width           =   4515
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   4
         Left            =   0
         TabIndex        =   7
         Top             =   3450
         Width           =   5190
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   540
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   5190
      End
   End
   Begin VB.Frame frmProgress 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1140
      Left            =   7500
      TabIndex        =   2
      Top             =   6000
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "VATBalanceSheet.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "VATBalanceSheet.frx":001C
         BarPictureMode  =   0
         BackPictureMode =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblMaster 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Τίτλος"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   150
         TabIndex        =   4
         Top             =   75
         Width           =   3765
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   2325
      Top             =   5025
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   2325
      Top             =   6225
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   5475
      Top             =   3375
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   12
      Left            =   0
      Top             =   2625
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   13
      Left            =   4425
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ισοζύγιο Φ.Π.Α."
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   30
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   720
      Left            =   225
      TabIndex        =   23
      Top             =   75
      Width           =   3690
   End
   Begin VB.Menu mnuHdrPopUp 
      Caption         =   "mnuHdrPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuΑποθήκευσηΠλάτουςΣτηλών 
         Caption         =   "Αποθήκευση πλάτους στηλών"
      End
   End
End
Attribute VB_Name = "VATBalanceSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnError As Boolean
Dim blnProcessing As Boolean

Function ColorizeVATBalance(curVATBalance)

        Select Case CCur(curVATBalance)
            Case 0
                lblLabel(0).Caption = "Μηδενικό υπόλοιπο"
                lblLabel(0).ForeColor = &H800000
                mskVATBalance.ForeColor = &H800000
            Case Is > 0
                lblLabel(0).Caption = "Πιστωτικό υπόλοιπο"
                lblLabel(0).ForeColor = &H8000&
                mskVATBalance.ForeColor = &H8000&
            Case Is < 0
                lblLabel(0).Caption = "Χρεωστικό υπόλοιπο"
                lblLabel(0).ForeColor = &H80&
                mskVATBalance.ForeColor = &HFF&
        End Select

End Function

Private Function FindRecordsAndPopulateBoxes()

    'Local μεταβλητές
    Dim curVATBalance As Currency
    
    'Μηδενίζω
    mskNetBuys.text = "0,00": mskVATBuys.text = "0,00"
    mskNetSales.text = "0,00": mskVATSales.text = "0,00"
    mskVATBalance.text = "0,00"
    
    'Χρωματίζω
    If ValidateFields Then
        If RefreshList(1) Then
            If RefreshList(2) Then
                curVATBalance = mskVATBuys.text - mskVATSales.text
                mskVATBalance.text = format(Abs(curVATBalance), "#,##0.00")
                ColorizeVATBalance curVATBalance
                DisableFields mskIssueFrom, mskIssueTo
                UpdateButtons Me, 2, 0, 1, 0
                blnProcessing = False
            End If
        End If
    End If
    
End Function

Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
        Case 0
            If ValidateFields Then FindRecordsAndPopulateBoxes
        Case 1
            AbortProcedure False
        Case 2
            AbortProcedure True
    End Select
    
End Sub

Private Function ValidateFields()

    ValidateFields = False
    
    'Από
    If DisplayMessage(1, 4, 1, "", mskIssueFrom.text) Then mskIssueFrom.SetFocus: Exit Function
    
    'Εως
    If DisplayMessage(1, 4, 1, "", mskIssueTo.text) Then mskIssueTo.SetFocus: Exit Function
    
    'Από <= Εως
    If DisplayMessage(14, 4, 1, "", mskIssueFrom.text, mskIssueTo.text) Then mskIssueFrom.SetFocus: Exit Function
    
    ValidateFields = True

End Function

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function
    
    If Not blnStatus Then
        lblLabel(0).Caption = "Υπόλοιπο"
        ColorizeControls Me
        ClearFields mskIssueFrom, mskIssueTo, mskNetBuys, mskVATBuys, mskNetSales, mskVATSales, mskVATBalance
        EnableFields mskIssueFrom, mskIssueTo
        mskIssueFrom.SetFocus
        UpdateButtons Me, 2, 1, 0, 1
    End If
    
    If blnStatus Then
        Unload Me
    End If

End Function

Private Function RefreshList(myReferToID)

    On Error GoTo ErrTrap
    
    'SQL
    Dim intIndex As Byte
    Dim strThisQuery As String
    Dim strParameters As String
    Dim strParFields As String
    Dim strThisParameter As String
    Dim strOrder As String
    Dim strLogic As String
    Dim arrQuery() As Variant
    Dim strSQL As String
    
    'Recordsets
    Dim rstRecordset As Recordset

    'Local μεταβλητές
    Dim curNetBuys As Currency
    Dim curVATBuys As Currency
    Dim curNetSales As Currency
    Dim curVATSales As Currency
    Dim curVATBalance As Currency
    
    'Αρχικές τιμές
    curNetBuys = 0
    curVATBuys = 0
    curNetSales = 0
    curVATSales = 0
    curVATBalance = 0
    
    intIndex = 0
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Αγορές - Πωλήσεις
    strSQL = "SELECT InvoiceIssueDate, Codes.CodeRefersTo, Codes.CodeSuppliers AS ColumnΑ, Codes.CodeCustomers AS ColumnB, InvoiceRestAmount, InvoiceExtraChargesAmount, InvoiceVATAmount " _
    & "FROM (Invoices " _
    & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID) "

    'Τύπος κίνησης
    strThisParameter = "intInvoice Integer"
    strThisQuery = "Invoices.InvoiceRefersToID = intInvoice"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = myReferToID
    
    'Από
    If IsDate(mskIssueFrom.text) Then
        strThisParameter = "datFrom Date"
        strThisQuery = "Invoices.InvoiceIssueDate >= datFrom"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskIssueFrom.text)
    End If
    
    'Εως
    If IsDate(mskIssueTo.text) Then
        strThisParameter = "datTo Date"
        strThisQuery = "Invoices.InvoiceIssueDate <= datTo"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskIssueTo.text)
    End If
        
    strOrder = " ORDER BY InvoiceIssueDate, InvoiceRefersToID"
        
    'Κριτήρια
    strParameters = "PARAMETERS " & strParameters & "; "
    strParFields = "WHERE " & strParFields
    strSQL = strParameters & strSQL & strParFields
    TempQuery.SQL = strSQL & strOrder
    For intIndex = 1 To UBound(arrQuery)
        TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
    Next intIndex
    
    'Ανοίγω το recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    'Αν δεν έχω εγγραφές, βγαίνω
    If rstRecordset.RecordCount = 0 Then blnError = False: RefreshList = True: Exit Function
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strAppTitle, rstRecordset
    
    'Προσωρινά
    UpdateButtons Me, 2, 0, 1, 0
    cmdButton(1).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Βρίσκω τις εγγραφές
    Do While Not rstRecordset.EOF
        UpdateProgressBar Me
        Select Case myReferToID
            'Αγορές
            Case 1
                'Αν αυξάνεται η αξία
                If rstRecordset![ColumnΑ] = "+" Then
                    curNetBuys = curNetBuys + rstRecordset![InvoiceRestAmount] + rstRecordset![InvoiceExtraChargesAmount]
                    curVATBuys = curVATBuys + rstRecordset![InvoiceVATAmount]
                End If
                'Αν μειώνεται η αξία
                If rstRecordset![ColumnΑ] = "-" Then
                    curNetBuys = curNetBuys - rstRecordset![InvoiceRestAmount] - rstRecordset![InvoiceExtraChargesAmount]
                    curVATBuys = curVATBuys - rstRecordset![InvoiceVATAmount]
                End If
            'Πωλήσεις
            Case 2
                'Αν αυξάνεται η αξία
                If rstRecordset![ColumnB] = "+" Then
                    curNetSales = curNetSales + rstRecordset![InvoiceRestAmount] + rstRecordset![InvoiceExtraChargesAmount]
                    curVATSales = curVATSales + rstRecordset![InvoiceVATAmount]
                End If
                'Αν μειώνεται η αξία
                If rstRecordset![ColumnB] = "-" Then
                    curNetSales = curNetSales - rstRecordset![InvoiceRestAmount] - rstRecordset![InvoiceExtraChargesAmount]
                    curVATSales = curVATSales - rstRecordset![InvoiceVATAmount]
                End If
        End Select
        rstRecordset.MoveNext
        DoEvents
        If Not blnProcessing Then Exit Do
    Loop
    
    'Ενημερώνω τα κουτάκια
    Select Case myReferToID
        Case 1
            mskNetBuys.text = format(curNetBuys, "#,##0.00")
            mskVATBuys.text = format(curVATBuys, "#,##0.00")
        Case 2
            mskNetSales.text = format(curNetSales, "#,##0.00")
            mskVATSales.text = format(curVATSales, "#,##0.00")
    End Select
    
    'Τελικές ενέργειες
    blnProcessing = False
    RefreshList = True
    cmdButton(1).Caption = "Νέα αναζήτηση"
    frmProgress.Visible = False
    
    Exit Function
    
UpdateSQLString:
    intIndex = intIndex + 1
    strParameters = IIf(intIndex > 1, strParameters & ", ", strParameters)
    strParFields = IIf(intIndex > 1, strParFields & strLogic, strParFields)
    strParameters = strParameters & strThisParameter
    strParFields = strParFields & strThisQuery
    ReDim Preserve arrQuery(intIndex)
    
    Return
    
ErrTrap:
    blnError = True
    cmdButton(1).Caption = "Νέα αναζήτηση"
    DisplayErrorMessage True, Err.Description
        
End Function

Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)

End Sub

Private Function CheckFunctionKeys(KeyCode, Shift)

    Dim CtrlDown
    
    CtrlDown = Shift + vbCtrlMask
    
    Select Case KeyCode
        Case vbKeyF10 And cmdButton(0).Enabled, vbKeyC And CtrlDown = 4 And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyEscape
            If cmdButton(1).Enabled Then cmdButton_Click 1: Exit Function
            If cmdButton(2).Enabled Then cmdButton_Click 2
        Case vbKeyF12
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    PositionControls Me, False: ColorizeControls Me
    ClearFields mskIssueFrom, mskIssueTo, mskNetBuys, mskVATBuys, mskNetSales, mskVATSales
    EnableFields mskIssueFrom, mskIssueTo
    UpdateButtons Me, 2, 1, 0, 1

End Sub

