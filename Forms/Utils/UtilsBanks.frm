VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form UtilsBanks 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   8970
   ClientLeft      =   15
   ClientTop       =   0
   ClientWidth     =   12975
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   12975
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   75
      TabIndex        =   9
      Top             =   7725
      Width           =   7515
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         Caption         =   "Δημιουργία"
         ButtonStyle     =   4
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
         BEVEL           =   2
         BEVELDEPTH      =   2
      End
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   1
         Left            =   1650
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         Caption         =   "Αποθήκευση"
         ButtonStyle     =   4
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
         Index           =   4
         Left            =   5925
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         Caption         =   "Κλείσιμο"
         ButtonStyle     =   4
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
         Index           =   2
         Left            =   3075
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         Caption         =   "Διαγραφή"
         ButtonStyle     =   4
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
         Index           =   3
         Left            =   4500
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         Caption         =   "Ακυρο"
         ButtonStyle     =   4
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
   Begin VB.Frame frmInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   75
      TabIndex        =   4
      Top             =   6225
      Width           =   4515
      Begin VB.TextBox txtCurrentGridRow 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
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
         Height          =   315
         Left            =   3675
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   450
         Width           =   780
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
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
         Height          =   315
         Left            =   75
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Grid.CurrentLine"
         Top             =   450
         Width           =   3540
      End
      Begin VB.TextBox txtBankID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
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
         Height          =   315
         Left            =   3675
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   75
         Width           =   780
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
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
         Height          =   315
         Left            =   75
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Banks.BankID"
         Top             =   75
         Width           =   3540
      End
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   75
         Top             =   825
         _ExtentX        =   953
         _ExtentY        =   953
         IconSizeX       =   26
         IconSizeY       =   32
         Size            =   14064
         Images          =   "UtilsBanks.frx":0000
         Version         =   131072
         KeyCount        =   4
         Keys            =   "ÿÿÿ"
      End
   End
   Begin iGrid300_10Tec.iGrid grdUtilsBanks 
      Height          =   6090
      Left            =   7125
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1125
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   10742
      Appearance      =   0
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
   End
   Begin UserControls.newText txtBankDescription 
      Height          =   465
      Left            =   1725
      TabIndex        =   1
      Top             =   1125
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   40
      Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   3825
      Top             =   8400
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   7350
      Top             =   7200
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   11775
      Top             =   3000
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   3
      Left            =   6675
      Top             =   1125
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   1
      Left            =   1275
      Top             =   1050
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   0
      Left            =   0
      Top             =   1125
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   4
      Left            =   7500
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Περιγραφή"
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
      TabIndex        =   3
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Τράπεζες"
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
      TabIndex        =   2
      Top             =   75
      Width           =   2115
   End
   Begin VB.Shape shpBackground 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   0
      Top             =   0
      Width           =   840
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   2
      Left            =   3300
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Menu mnuHdrPopUp 
      Caption         =   "mnuHdrPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuΑποθήκευσηΠλάτουςΣτηλών 
         Caption         =   "Αποθήκευση πλάτους στηλών"
      End
   End
End
Attribute VB_Name = "UtilsBanks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim blnStatus As Boolean

Private Function ValidateFields()

    ValidateFields = False
    
    'Περιγραφή
    If DisplayMessage(1, 4, 1, "", txtBankDescription.text) Then txtBankDescription.SetFocus: Exit Function
    
    ValidateFields = True

End Function

Private Function AbortProcedure(blnStatus)
    
    If Not blnStatus Then
        If MyMsgBox(3, strAppTitle, strMessages(3), 2) Then
            blnStatus = False
            ClearFields txtBankID, txtBankDescription
            DisableFields txtBankDescription
            grdUtilsBanks.SetFocus
            UpdateButtons Me, 4, 1, 0, 0, 0, 1
        End If
        Exit Function
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Function DeleteRecord()
    
    If MainDeleteRecord("CommonDB", "Banks", strAppTitle, "ID", txtBankID.text, "True") Then
        PopulateGrid
        HighlightNextRow grdUtilsBanks, Val(txtCurrentGridRow.text), 2, True
        ClearFields txtBankID, txtBankDescription
        DisableFields txtBankDescription
        UpdateButtons Me, 4, 1, 0, 0, 0, 1
    End If

End Function

Private Function NewRecord()
    
    blnStatus = True
    ClearFields txtBankID, txtBankDescription
    EnableFields txtBankDescription
    UpdateButtons Me, 4, 0, 1, 0, 1, 0
    txtBankDescription.SetFocus

End Function

Private Function SaveRecord()
    
    Dim blnNotError
    
    If Not ValidateFields Then Exit Function
    
    blnNotError = MainSaveRecord("CommonDB", "Banks", blnStatus, strAppTitle, "ID", txtBankID.text, txtBankDescription.text, 1, strCurrentUser)
        
    If IsNumeric(blnNotError) And blnNotError Then
        txtBankID.text = blnNotError
        PopulateGrid
        HighlightRow grdUtilsBanks, 1, txtBankID.text, True
        ClearFields txtBankID, txtBankDescription
        DisableFields txtBankDescription
        UpdateButtons Me, 4, 1, 0, 0, 0, 1
    End If
    
End Function

Private Function SeekRecord()

    Dim blnEnableDelete As Boolean
    
    If grdUtilsBanks.RowCount = 0 Then Exit Function
    
    ClearFields txtBankID, txtBankDescription
    DisableFields txtBankDescription
    
    blnEnableDelete = SimpleSeek("Checks", "BankID", grdUtilsBanks.CellValue(grdUtilsBanks.CurRow, 1))
    
    If MainSeekRecord("CommonDB", "Banks", "ID", grdUtilsBanks.CellValue(grdUtilsBanks.CurRow, 1), True, txtBankID, txtBankDescription) Then
        blnStatus = False
        EnableFields txtBankDescription
        UpdateButtons Me, 4, 0, 1, IIf(blnEnableDelete, 1, 0), 1, 0
        txtBankDescription.SetFocus
        txtCurrentGridRow.text = grdUtilsBanks.CurRow
    End If
    
End Function

Private Sub cmdButton_Click(Index As Integer)
                                                                
    Select Case Index
        Case 0
            NewRecord
        Case 1
            SaveRecord
        Case 2
            DeleteRecord
        Case 3
            AbortProcedure False
        Case 4
            AbortProcedure True
    End Select

End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdUtilsBanks, 25, GetSetting(strAppTitle, "Layout Strings", "grdUtilsBanks"), "04NCNID,40NLNDescription", "ID,Περιγραφή"
        Me.Refresh
        PopulateGrid
    End If
    
    'AddDummyLines grdUtilsBanks, "99999", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAΑΑΑΑΑΑΑΑΑΑ"

End Sub

Private Function PopulateGrid()

    If FillGridFromDB("CommonDB", grdUtilsBanks, "Banks", "", "", "", 2, 0, 1) Then
        grdUtilsBanks.SetFocus
        grdUtilsBanks.SetCurCell 1, 1
    End If

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)
    
End Sub

Private Function CheckFunctionKeys(KeyCode, Shift)
    
    Dim CtrlDown
    
    CtrlDown = Shift + vbCtrlMask
    
    Select Case KeyCode
        Case vbKeyInsert And cmdButton(0).Enabled, vbKeyN And CtrlDown = 4 And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyF10 And cmdButton(1).Enabled, vbKeyS And CtrlDown = 4 And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyF3 And cmdButton(2).Enabled, vbKeyD And CtrlDown = 4 And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyEscape
            If cmdButton(3).Enabled Then cmdButton_Click 3: Exit Function
            If cmdButton(4).Enabled Then cmdButton_Click 4
        Case vbKeyF12
            ToggleInfoPanel Me
End Select

End Function

Private Sub Form_Load()
    
    SetUpGrid lstIconList, grdUtilsBanks
    PositionControls Me, False: ColorizeControls Me
    ClearFields txtBankID, txtBankDescription
    DisableFields txtBankDescription
    UpdateButtons Me, 4, 1, 0, 0, 0, 1

End Sub

Private Sub grdUtilsBanks_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    SeekRecord

End Sub

Private Sub grdUtilsBanks_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdUtilsBanks_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SeekRecord
    
End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdUtilsBanks", grdUtilsBanks.LayoutCol

End Sub

