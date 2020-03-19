VERSION 5.00
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form CommonLogin 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Wheels and Tires"
   ClientHeight    =   6690
   ClientLeft      =   0
   ClientTop       =   1365
   ClientWidth     =   9495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000011&
   Icon            =   "CommonLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboCompanies 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   5250
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4275
      Width           =   3630
   End
   Begin VB.ComboBox cboUsers 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   6375
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3825
      Width           =   2505
   End
   Begin UserControls.newText txtPassword 
      Height          =   465
      Left            =   7350
      TabIndex        =   2
      Top             =   4725
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   0
      MaxLength       =   10
      PasswordChar    =   "*"
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
   Begin GurhanButtonOCX.GurhanButton cmdButton 
      Height          =   690
      Index           =   0
      Left            =   6075
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1217
      Caption         =   "Εναρξη"
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
      MousePointer    =   2
      ShowFocusRect   =   0   'False
      XPColor_Hover   =   8438015
      BackColor       =   8438015
      ForeColor       =   0
   End
   Begin GurhanButtonOCX.GurhanButton cmdButton 
      Height          =   690
      Index           =   1
      Left            =   7500
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5475
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
      MousePointer    =   2
      ShowFocusRect   =   0   'False
      BackColor       =   8421631
      ForeColor       =   0
   End
   Begin VB.Label lblCopyright 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "(c) John Sourvinos 1996-2017"
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
      Left            =   5475
      TabIndex        =   8
      Top             =   3075
      Width           =   3390
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Εταιρία"
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
      Height          =   255
      Index           =   1
      Left            =   4650
      TabIndex        =   7
      Top             =   4335
      Width           =   510
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Πλατφόρμα: Win32"
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
      Index           =   99
      Left            =   5475
      TabIndex        =   6
      Top             =   2775
      Width           =   3390
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Χρήστης"
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
      Height          =   255
      Index           =   0
      Left            =   5700
      TabIndex        =   5
      Top             =   3870
      Width           =   585
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Κωδικός"
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
      Height          =   255
      Index           =   2
      Left            =   6675
      TabIndex        =   4
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblProgress 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Πρόοδος εργασιών"
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
      Height          =   255
      Left            =   3375
      TabIndex        =   3
      Top             =   3375
      Width           =   5505
   End
   Begin VB.Image imgImage 
      Height          =   6540
      Left            =   75
      Picture         =   "CommonLogin.frx":1CCA
      Stretch         =   -1  'True
      Top             =   75
      Width           =   9315
   End
End
Attribute VB_Name = "CommonLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function CheckForEAFDSS()

    'Ελεγχος
    If blnCheckEAFDSS Then
        If Not (IsProcessRunning(strEAFDSS) = True) Then
            If MyMsgBox(4, strAppTitle, strMessages(28), 2) Then
                CheckForEAFDSS = True
            Else
                CheckForEAFDSS = False
            End If
        Else
            CheckForEAFDSS = True
        End If
    Else
        CheckForEAFDSS = True
    End If

End Function

Private Function LoadUsers()
    
    On Error GoTo ErrTrap
    
    Dim rsUsers As Recordset
    
    Set UsersDB = DBEngine.OpenDataBase(strPathName + "Users.mdb", False, False)
    
    Set rsUsers = UsersDB.OpenRecordset("Users")
    With rsUsers
        While Not .EOF
            cboUsers.AddItem !username
            .MoveNext
        Wend
    End With
    
    LoadUsers = True
    UsersDB.Close
    
    Exit Function
    
ErrTrap:
    Me.Tag = "False"
    LoadUsers = False
    DisplayErrorMessage True, Err.Description

End Function

Private Function CheckFunctionKeys(KeyCode, Shift)
    
    Dim CtrlDown
    
    CtrlDown = Shift + vbCtrlMask
    
    Select Case KeyCode
        Case vbKeyF10 And cmdButton(0).Enabled, vbKeyC And CtrlDown = 4 And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyEscape And cmdButton(1).Enabled
            cmdButton_Click 1
    End Select

End Function

Private Function CloseApp()

    CloseApp = False
    
    If MyMsgBox(2, strAppTitle, strMessages(19), 2) Then
        CloseApp = True
    End If

End Function

Private Function LoadDatabaseSettings()

    strDatabaseType = GetSetting(appName:=strAppTitle, Section:="Databases", Key:="Type")
    strDatabaseName = GetSetting(appName:=strAppTitle, Section:="Databases", Key:="Name")
    strDatabasePort = GetSetting(appName:=strAppTitle, Section:="Databases", Key:="Port")
    strDatabaseServer = GetSetting(appName:=strAppTitle, Section:="Databases", Key:="Server")
    
End Function

Private Function LoadCompanies()
    
    On Error GoTo ErrTrap
    
    Dim strCompanies As String
    Dim strCompany As String
    Dim bytPosition As Byte
    Dim obj As Object
    
    cboCompanies.Clear
    
    If strDatabaseType = "Access" Then
        strPathName = GetSetting(appName:=strAppTitle, Section:="Path Names", Key:="Database Path Name")
        strCompanies = Dir(strPathName & "*.mdb")
        If strCompanies = "" Then Exit Function
        Do While strCompanies <> ""
            If strCompanies <> "Printers.mdb" And strCompanies <> "Users.mdb" Then
                strCompany = ""
                bytPosition = 1
                While Mid(strCompanies, bytPosition, 1) <> "."
                    strCompany = strCompany + Mid(strCompanies, bytPosition, 1)
                    bytPosition = bytPosition + 1
                Wend
                cboCompanies.AddItem strCompany
            End If
            strCompanies = Dir
        Loop
    End If
    
    LoadCompanies = True
    
    Exit Function
    
ErrTrap:
    Me.Tag = "False"
    LoadCompanies = False
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function Start()

    lblProgress.Caption = strMessages(27)
    lblProgress.Refresh
    
    If App.PrevInstance Then
        If MyMsgBox(4, strAppTitle, strMessages(16), 1) Then
        End If
        CloseApp
        End
    End If
    
    strCompanyName = cboCompanies.text & ".mdb"
    strCurrentUser = cboUsers.text
    
    If Not IsCorrectPassword(cboUsers.text, txtPassword.text) Then
        If MyMsgBox(4, strAppTitle, strMessages(15), 1) Then
        End If
        ClearFields lblProgress
        Exit Function
    End If
    
    If OpenDataBase(strCompanyName) Then
        If LoadSettings Then
            If Not CheckForEAFDSS Then End
            CommonMain.Caption = "Server: " & strPathName & " - Εταιρία: " & Left(strCompanyName, Len(strCompanyName) - 4) & " - Χρήστης: " & strCurrentUser
            If Not CommonMain.Visible Then
                blnAppIsRunning = True
                CommonMain.Show
            End If
            Unload Me
        Else
            ClearFields lblProgress
        End If
    End If

End Function

Private Sub cboCompanies_KeyPress(KeyAscii As Integer)

    ValidateInput (KeyAscii)

End Sub

Private Sub cboUsers_KeyPress(KeyAscii As Integer)

    ValidateInput (KeyAscii)

End Sub

Private Sub cmdButton_Click(Index As Integer)

    Dim obj As Object
    
    Select Case Index
        Case 0
            If ValidateFields Then Start
        Case 1
            If blnAppIsRunning Then
                Unload Me
            Else
                If CloseApp Then
                    For Each obj In Forms
                        Unload obj
                    Next
                End If
            End If
    End Select
        
End Sub

Private Sub cmdButton_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    CheckForArrows (KeyCode)

End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        Me.Tag = "False"
        ClearFields lblProgress
        cboUsers.SetFocus
        'cboUsers.ListIndex = 1
        'cboCompanies.ListIndex = 0
        'txtPassword.text = "1701"
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)
    
End Sub

Private Sub Form_Load()

    strReportsPathName = GetSetting(appName:=strAppTitle, Section:="Path Names", Key:="Reports Path Name")
    
    Me.Tag = "True"
    Me.Show
    lblCopyright.Caption = "(c) John Sourvinos 1996-" & Year(Date)
    ClearFields lblProgress, cboUsers, cboCompanies, txtPassword
    strApplicationEXEName = GetSetting(strAppTitle, "Settings", "Application EXE Name")
    LoadMessages
    LoadDatabaseSettings
    If Not LoadCompanies Then Exit Sub
    If Not LoadUsers Then Exit Sub
    If GetSetting(appName:=strAppTitle, Section:="Settings", Key:="IsDevelopment") = "1" Then
        cboUsers.ListIndex = 0
        cboCompanies.ListIndex = 0
        txtPassword.text = "1701"
    End If

End Sub
    
Private Function ValidateFields()

    ValidateFields = False
    
    'Χρήστες
    If DisplayMessage(1, 4, 1, "", cboUsers.text) Then cboUsers.SetFocus: Exit Function
    
    'Εταιρία
    If DisplayMessage(1, 4, 1, "", cboCompanies.text) Then cboCompanies.SetFocus: Exit Function
    
    'Κωδικός
    If DisplayMessage(1, 4, 1, "", txtPassword.text) Then txtPassword.SetFocus: Exit Function
    
    ValidateFields = True

End Function


