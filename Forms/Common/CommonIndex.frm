VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form CommonIndex 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8415
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4575
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkShowInactiveRecords 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "Εμφάνιση ανενεργών εγγραφών"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   315
      Left            =   -450
      TabIndex        =   5
      Top             =   300
      Width           =   3015
   End
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   75
      TabIndex        =   2
      Top             =   7350
      Width           =   4365
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   75
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   75
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         Caption         =   "Επιλογή"
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
         MousePointer    =   2
         ShowFocusRect   =   0   'False
         BackColor       =   16777152
      End
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   2
         Left            =   2925
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   75
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
         MousePointer    =   2
         ShowFocusRect   =   0   'False
         BackColor       =   8421631
      End
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   1
         Left            =   1500
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   75
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         Caption         =   "Καρτέλα"
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
         MousePointer    =   2
         ShowFocusRect   =   0   'False
         BackColor       =   12640511
      End
   End
   Begin iGrid300_10Tec.iGrid grdGrid 
      Height          =   6165
      Left            =   300
      TabIndex        =   0
      Top             =   900
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   10874
      Appearance      =   0
      BackColor       =   12648447
      BorderStyle     =   1
      DefaultRowHeight=   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483631
      ScrollBarStyle  =   2
   End
   Begin vbalIml6.vbalImageList lstIconList 
      Left            =   375
      Top             =   6450
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   4592
      Images          =   "CommonIndex.frx":0000
      Version         =   131072
      KeyCount        =   4
      Keys            =   "ÿÿÿ"
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ευρετήριο"
      BeginProperty Font 
         Name            =   "Aka-Acid-Steelfish"
         Size            =   26.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   630
      Left            =   300
      TabIndex        =   1
      Top             =   75
      Width           =   1470
   End
   Begin VB.Shape shpShape 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   6315
      Left            =   225
      Top             =   825
      Width           =   2565
   End
End
Attribute VB_Name = "CommonIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function ShowItemLedger()

    Dim lngCurrentRow As Long
    
    lngCurrentRow = grdGrid.CurRow
    
    With ItemsLedger
        .txtCategoryID.text = grdGrid.CellValue(lngCurrentRow, 2)
        .txtCategoryShortDescription.text = grdGrid.CellValue(lngCurrentRow, 7)
        .lblCategoryDescription.Caption = grdGrid.CellValue(lngCurrentRow, 4)
        .txtManufacturerID.text = grdGrid.CellValue(lngCurrentRow, 3)
        .txtManufacturerDescription.text = grdGrid.CellValue(lngCurrentRow, 6)
        .txtItemID.text = grdGrid.CellValue(lngCurrentRow, 1)
        .txtItemDescription.text = grdGrid.CellValue(lngCurrentRow, 5)
        .txtTable.text = "Items"
        .Tag = "True"
        DisableFields .txtCategoryShortDescription, .txtManufacturerDescription, .txtItemDescription, .cmdIndex(0), .cmdIndex(1), .cmdIndex(2)
        .Show 1, Me
    End With
    
End Function

Private Sub chkShowInactiveRecords_Click()

    ToggleInactiveRecords

End Sub

Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
        Case 0
            Me.Hide
        Case 1
            ShowItemLedger
        Case 2
            AbortProcedure
    End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)

End Sub

Private Function AbortProcedure()
    
    On Error GoTo ErrTrap
    
    Dim lngCol As Long
    
    If cmdButton(1).Enabled Then
        For lngCol = 1 To grdGrid.colCount
            grdGrid.CellValue(CommonIndex.grdGrid.CurRow, lngCol) = ""
        Next lngCol
    End If
    
    Me.Hide
    
    Exit Function
    
ErrTrap:
    Me.Hide
    Exit Function
    
End Function

Private Function CheckFunctionKeys(KeyCode, Shift)
    
    Dim CtrlDown
    
    CtrlDown = Shift + vbCtrlMask
    
    Select Case KeyCode
        Case vbKeyReturn
            cmdButton_Click 0
        Case vbKeyF4 And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyEscape
            cmdButton_Click 2
        Case vbKeyA And CtrlDown = 4 And chkShowInactiveRecords.Visible
            chkShowInactiveRecords.Value = IIf(chkShowInactiveRecords.Value = 0, 1, 0)
    End Select
    
End Function

Private Sub Form_Load()

    SetUpGrid lstIconList, grdGrid
    ColorizeGrid grdGrid
    
End Sub

Private Sub grdGrid_ColHeaderMouseEnter(ByVal lCol As Long)

    grdGrid.Header.Buttons = True

End Sub

Private Sub grdGrid_ColHeaderMouseLeave(ByVal lCol As Long)

    grdGrid.Header.Buttons = False

End Sub

Private Sub grdGrid_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    Me.Hide

End Sub

Private Sub grdGrid_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeyF4 And cmdButton(1).Enabled Then ShowItemLedger

End Sub

Private Sub grdGrid_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then Me.Hide

End Sub

Private Function ToggleInactiveRecords()

    Dim lngRow As Long
    
    For lngRow = 1 To grdGrid.RowCount
        If grdGrid.CellFont(lngRow, 1).Italic Then
            grdGrid.RowVisible(lngRow) = Not grdGrid.RowVisible(lngRow)
        End If
    Next lngRow

End Function

