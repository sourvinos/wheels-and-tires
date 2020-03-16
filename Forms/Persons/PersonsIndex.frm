VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form PersonsIndex 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   9765
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   17970
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   17970
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmProgress 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1140
      Left            =   8475
      TabIndex        =   27
      Top             =   7650
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "PersonsIndex.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "PersonsIndex.frx":001C
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
         TabIndex        =   29
         Top             =   75
         Width           =   3765
      End
   End
   Begin VB.Frame frmContainer 
      BackColor       =   &H00008080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   9615
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   17565
      Begin VB.Frame frmButtonFrame 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   75
         TabIndex        =   18
         Top             =   8850
         Width           =   11790
         Begin GurhanButtonOCX.GurhanButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            Caption         =   "Συνέχεια"
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
            Index           =   1
            Left            =   1650
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            Caption         =   "Επεξεργασία"
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
            Index           =   7
            Left            =   10200
            TabIndex        =   21
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
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            Caption         =   "Εκτύπωση"
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
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            Caption         =   "Μαζική επεξεργασία"
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
            Index           =   5
            Left            =   7350
            TabIndex        =   24
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
            Index           =   6
            Left            =   8775
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            Caption         =   "Νέα αναζήτηση"
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
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            Caption         =   "Δημιουργία αρχείου PDF"
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
         Height          =   1815
         Left            =   8400
         TabIndex        =   11
         Top             =   5700
         Width           =   4515
         Begin VB.TextBox OppositeTable 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            TabIndex        =   33
            TabStop         =   0   'False
            Text            =   "OppositeTable"
            Top             =   450
            Width           =   1965
         End
         Begin VB.TextBox txtOppositeTable 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            Left            =   2100
            TabIndex        =   32
            TabStop         =   0   'False
            Text            =   "3"
            Top             =   450
            Width           =   2340
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "RefersTo"
            Top             =   825
            Width           =   1965
         End
         Begin VB.TextBox txtRefersTo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            Left            =   2100
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   "6"
            Top             =   825
            Width           =   2340
         End
         Begin VB.TextBox Text20 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "Table"
            Top             =   75
            Width           =   1965
         End
         Begin VB.TextBox txtTable 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            Left            =   2100
            TabIndex        =   12
            TabStop         =   0   'False
            Text            =   "1"
            Top             =   75
            Width           =   2340
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   1200
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   2296
            Images          =   "PersonsIndex.frx":0038
            Version         =   131072
            KeyCount        =   2
            Keys            =   "ÿ"
         End
      End
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   2565
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   6150
         Width           =   8190
         Begin UserControls.newText txtDescription 
            Height          =   465
            Left            =   1575
            TabIndex        =   2
            Top             =   825
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
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
         Begin UserControls.newText txtTaxNo 
            Height          =   465
            Left            =   1575
            TabIndex        =   3
            Top             =   1350
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   820
            Alignment       =   2
            ForeColor       =   0
            MaxLength       =   10
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
            Left            =   75
            TabIndex        =   31
            Top             =   75
            Width           =   1665
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
            Left            =   4725
            TabIndex        =   30
            Top             =   75
            Width           =   3315
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   1
            Left            =   1125
            Top             =   825
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
            Left            =   7725
            Top             =   600
            Visible         =   0   'False
            Width           =   465
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
         Begin VB.Label Label1 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
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
            TabIndex        =   10
            Top             =   2100
            Width           =   8190
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Α.Φ.Μ."
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
            TabIndex        =   5
            Top             =   1425
            Width           =   690
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Επωνυμία"
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
            TabIndex        =   4
            Top             =   900
            Width           =   690
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
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
            Width           =   8190
         End
      End
      Begin iGrid300_10Tec.iGrid grdPersonsIndex 
         Height          =   7290
         Left            =   75
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1500
         Width           =   17415
         _ExtentX        =   30718
         _ExtentY        =   12859
         Appearance      =   0
         BackColor       =   14737632
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
      Begin VB.Label lblSelected 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Επιλεγμένες 0 εγγραφές"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   315
         Left            =   2550
         TabIndex        =   15
         Top             =   825
         Width           =   14940
      End
      Begin VB.Label lblRecordCount 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Βρέθηκαν 0 εγγραφές"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   315
         Left            =   75
         TabIndex        =   14
         Top             =   1125
         Width           =   6690
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ευρετήριο συναλλασόμενων"
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
         Left            =   75
         TabIndex        =   8
         Top             =   75
         Width           =   6585
      End
      Begin VB.Label lblCriteria 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Κριτήρια αναζήτησης"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   315
         Left            =   2550
         TabIndex        =   7
         Top             =   1125
         Width           =   14940
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
   End
   Begin VB.Menu mnuHdrPopUp 
      Caption         =   "mnuHdrPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuΑποθήκευσηΠλάτουςΣτηλών 
         Caption         =   "Αποθήκευση πλάτους στηλών"
      End
   End
End
Attribute VB_Name = "PersonsIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngRowCount As Long
Dim blnError As Boolean
Dim blnProcessing As Boolean
Dim blnBatchProcessing As Boolean

Private Function EditRecord()

    If Persons.SeekRecord(grdPersonsIndex.CellValue(grdPersonsIndex.CurRow, "ID"), txtTable.text, txtRefersTo.text) Then
        If Persons.Visible Then
            Unload Me
        Else
            With Persons
                .Tag = "True"
                .txtTable.text = txtTable.text
                .txtOppositeTable.text = txtOppositeTable.text
                .txtRefersTo.text = Val(txtRefersTo.text)
                .lblTitle.Caption = IIf(txtRefersTo.text = "3", "Προμηθευτές", "Πελάτες")
                .Show 1
            End With
        End If
    End If

End Function

Private Function EnableBatchProcess()

    blnBatchProcessing = True
    UpdateButtons Me, 7, 0, 0, 0, 0, 0, 1, 1, 0
    cmdButton(6).Caption = "Ακυρο"
    EnableGrid grdPersonsIndex, True, grdPersonsIndex.CurRow, 2

End Function

Private Function FindRecordsAndPopulateGrid()

    If RefreshList > 0 Then
        UpdateRecordCount lblRecordCount, lngRowCount
        UpdateCriteriaLabels txtDescription.text, txtTaxNo.text
        EnableGrid grdPersonsIndex, False
        HighlightRow grdPersonsIndex, 1, "", True
        UpdateButtons Me, 7, 0, 1, 1, 1, 1, 0, 1, 0
    Else
        UpdateButtons Me, 7, 1, 0, 0, 0, 0, 0, 0, 1
        If Not blnError Then
            If blnProcessing Then
                If MyMsgBox(1, strAppTitle, strMessages(23), 1) Then
                End If
            Else
                If MyMsgBox(1, strAppTitle, strMessages(8), 1) Then
                End If
            End If
        End If
        blnError = False
        blnProcessing = False
        frmCriteria(0).Visible = True
        txtDescription.SetFocus
    End If
    
End Function

Function CreateUnicodeFile(myPrinterType, myEAFDSSString, myInvoiceHeight, myDetailLines, myTopMargin, myLeftMargin)

    On Error GoTo ErrTrap
    
    Dim lngRow As Long
    Dim intProcessedDetailLines As Integer
    Dim intPageNo As Integer
    
    intPageNo = 0
    intProcessedDetailLines = 0
    
    Open strUnicodeFile For Output As #1
    InitReport myPrinterType, myEAFDSSString, myInvoiceHeight
    GoSub Headers
    
    With grdPersonsIndex
        For lngRow = 1 To .RowCount
            Print #1, Tab(1); Left(.CellText(lngRow, "Description"), 50); Tab(52); Left(.CellText(lngRow, "TaxNo"), 10); Tab(62); Left(.CellText(lngRow, "Profession"), 30); Tab(93); Left(.CellText(lngRow, "Phones"), 44)
            
            intProcessedDetailLines = intProcessedDetailLines + 1
            If intProcessedDetailLines > myDetailLines Then
                If lngRow < .RowCount Then
                    Print #1, ""
                    Print #1, strMessages(24)
                    GoSub Headers
                    Print #1, strMessages(13)
                    Print #1, ""
                    intProcessedDetailLines = intProcessedDetailLines + 2
                End If
            End If
        Next lngRow
    End With
    
    Print #1, ""
    Print #1, strMessages(25)
    
    Close #1
    
    CreateUnicodeFile = strUnicodeFile
    
    Exit Function
    
Headers:
    intPageNo = intPageNo + 1
    PrintHeadings 136, intPageNo, CustomUpperCase(lblTitle.Caption), CustomUpperCase(lblCriteria.Caption), "", myTopMargin
    PrintColumnHeadings 1, "ΕΠΩΝΥΜΙΑ", 52, "Α.Φ.Μ.", 62, "ΔΡΑΣΤΗΡΙΟΤΗΤΑ", 93, "ΤΗΛΕΦΩΝΑ"
    Print #1, ""
    intProcessedDetailLines = 6
    
    Return
    
ErrTrap:
    Close #1
    CreateUnicodeFile = "Error"
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function CheckPersonForTransactions(myTable, myRefersTo, myID)

    Dim strSQL As String
    Dim rstRecordset As Recordset
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    strSQL = "SELECT InvoiceID FROM Invoices WHERE InvoicePersonID = " & myID & " AND (InvoiceRefersToID = " & myRefersTo & " OR InvoiceRefersToID = " & myRefersTo - 2 & ")"
    
    TempQuery.SQL = strSQL
    
    Set rstRecordset = TempQuery.OpenRecordset()
    
    If rstRecordset.RecordCount >= 1 Then
        CheckPersonForTransactions = "Y"
    Else
        CheckPersonForTransactions = ""
    End If

End Function

Private Function UpdateCriteriaLabels(Description, TaxNo)

    Dim strCriteriaA As String
    Dim strCriteriaB As String

    If Description = "" Then
        strCriteriaA = "Επωνυμία [ ΟΛΟΙ ]"
    Else
        If Left(Description, 1) <> "*" Then strCriteriaA = "Επωνυμία αρχίζει από [ " & UCase(Description) & " ]"
        If Left(Description, 1) = "*" Then strCriteriaA = "Επωνυμία περιέχει το [ " & UCase(Right(Description, Len(Description) - 1)) & " ]"
    End If
    
    If TaxNo = "" Then strCriteriaB = "Α.Φ.Μ. [ ΟΛΟΙ ]"
    If TaxNo <> "" Then strCriteriaB = "Α.Φ.Μ. αρχίζει από [ " & TaxNo & " ]"
    
    lblCriteria.Caption = strCriteriaA & " " & strCriteriaB
    
End Function

Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
        Case 0
            FindRecordsAndPopulateGrid
        Case 1
            EditRecord
        Case 2
            PrintRecords Me, "Print", False, "PrinterPrintsReportsID"
        Case 3
            PrintRecords Me, "CreatePDF", True, "PrinterPrintsReportsID"
        Case 4
            EnableBatchProcess
        Case 5
            SaveRecords grdPersonsIndex.CurRow
        Case 6
            AbortProcedure False
        Case 7
            AbortProcedure True
    End Select
    
End Sub

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function
    
    If Not blnStatus Then
        If Not blnBatchProcessing Then
            ClearFields grdPersonsIndex, lblRecordCount, lblCriteria, lblSelected
            frmCriteria(0).Visible = True
            txtDescription.SetFocus
            UpdateButtons Me, 7, 1, 0, 0, 0, 0, 0, 0, 1
        Else
            If MyMsgBox(3, strAppTitle, strMessages(3), 2) Then
                EnableGrid grdPersonsIndex, False, grdPersonsIndex.CurRow
                UpdateButtons Me, 7, 0, 1, 1, 1, 1, 0, 1, 0
                blnBatchProcessing = False
            End If
        End If
    End If
    
    If blnStatus Then
        Unload Me
    End If

End Function

Private Function RefreshList()

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

    'Αρχικές τιμές
    intIndex = 0
    lngRowCount = 0
    frmCriteria(0).Visible = False
    blnBatchProcessing = False
    Set TempQuery = CommonDB.CreateQueryDef("")

    'Πλέγμα
    With grdPersonsIndex
        .Clear
        .Editable = False
        .Redraw = False
        .RowMode = False
    End With
    
    'Κυρίως διαδικασία
    strSQL = "SELECT ID, Description, TaxNo, " & txtTable.text & ".TaxOfficeID, Profession, Address, City, Phones, InCharge, VATStateID, Email, TaxOfficeDescription, CountryID " _
    & "FROM " & txtTable.text & " " _
    & "LEFT JOIN TaxOffices ON " & txtTable.text & ".TaxOfficeID = TaxOffices.TaxOfficeID "
    
    'Επωνυμία
    If txtDescription.text <> "" Then
        strThisParameter = "strDescription String"
        If Left(txtDescription.text, 1) <> "*" Then
            strThisQuery = "Left(Description, Len(strDescription))= strDescription"
        End If
        If Left(txtDescription.text, 1) = "*" Then
            strThisQuery = "InStr(Description, strDescription)"
        End If
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = IIf(Left(txtDescription.text, 1) <> "*", txtDescription.text, Right(txtDescription.text, Len(txtDescription.text) - 1))
    End If
    
    'Α.Φ.Μ.
    If txtTaxNo.text <> "" Then
        strThisParameter = "strTaxNo String"
        strThisQuery = "Left(TaxNo, Len(strTaxNo)) = strTaxNo"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = txtTaxNo.text
    End If

    'Ταξινόμηση
    strOrder = " ORDER BY Description"
    
    'Προσθέτω τα κριτήρια
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
    End If
    
    TempQuery.SQL = strSQL & strOrder
    
    'Κριτήρια
    If strThisParameter <> "" Then
        For intIndex = 1 To UBound(arrQuery)
            TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
        Next intIndex
    End If
    
    'Ανοίγω το recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    'Αν δεν έχω εγγραφές, βγαίνω
    If rstRecordset.RecordCount = 0 Then blnError = False: RefreshList = False: Exit Function
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strAppTitle, rstRecordset
    
    'Προσωρινά
    UpdateButtons Me, 7, 0, 0, 0, 0, 0, 0, 1, 0
    cmdButton(6).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        grdPersonsIndex.AddRow , , , , , , , rstRecordset.RecordCount
        Do While Not .EOF
            lngRowCount = lngRowCount + 1
            grdPersonsIndex.CellValue(lngRowCount, "ID") = !ID
            grdPersonsIndex.CellValue(lngRowCount, "Description") = !Description
            grdPersonsIndex.CellValue(lngRowCount, "TaxNo") = !TaxNo
            grdPersonsIndex.CellValue(lngRowCount, "TaxOfficeID") = !TaxOfficeID
            grdPersonsIndex.CellValue(lngRowCount, "Profession") = !Profession
            grdPersonsIndex.CellValue(lngRowCount, "Address") = !Address
            grdPersonsIndex.CellValue(lngRowCount, "City") = !City
            grdPersonsIndex.CellValue(lngRowCount, "Phones") = !Phones
            grdPersonsIndex.CellValue(lngRowCount, "InCharge") = !InCharge
            grdPersonsIndex.CellValue(lngRowCount, "VATStateID") = !VATStateID
            grdPersonsIndex.CellValue(lngRowCount, "Email") = !Email
            grdPersonsIndex.CellValue(lngRowCount, "TaxOfficeDescription") = !TaxOfficeDescription
            grdPersonsIndex.CellValue(lngRowCount, "CountryID") = !CountryID
            grdPersonsIndex.CellValue(lngRowCount, "HasTransactions") = CheckPersonForTransactions(txtTable.text, txtRefersTo.text, !ID)
            UpdateProgressBar Me
            .MoveNext
            DoEvents
            If Not blnProcessing Then Exit Do
        Loop
    End With
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        RefreshList = 0
        ClearFields grdPersonsIndex
    Else
        RefreshList = rstRecordset.RecordCount
        blnProcessing = False
    End If
    
    'Τελικές ενέργειες
    cmdButton(6).Caption = "Νέα αναζήτηση"
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
    ClearFields grdPersonsIndex, frmProgress
    cmdButton(6).Caption = "Νέα αναζήτηση"
    DisplayErrorMessage True, Err.Description
    
End Function

Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdPersonsIndex, 44, GetSetting(strAppTitle, "Layout Strings", "grdPersonsIndex"), _
            "05NCNID,50NLNDescription,15NCNTaxNo,05NCNXTaxOfficeID,50NLNProfession,50NLNAddress,50NLNCity,50NLNPhones,50NLNInCharge,40NCNXVATStateID,40NLNEmail,03NCNSelected,40NLNTaxOfficeDescription,05NCNCountryID,03NCNHasTransactions", _
            "ID,Επωνυμία,Α.Φ.Μ.,Οικονομική υπηρεσία,Δραστηριότητα,Διεύθυνση,Πόλη,Τηλέφωνα,Υπεύθυνος,Καθεστώς Φ.Π.Α.,E-mail,Ε,Οικονομική υπηρεσία,Χώρα,Κ"
        Me.Refresh
        frmCriteria(0).Visible = True
        txtDescription.SetFocus
    End If
    
    'AddDummyLines grdPersonsIndex, 5, 50, 15, 5, 50, 40, 40, 30, 40, 5, 40, 5, 40, 4

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
        Case vbKeyE And CtrlDown = 4 And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyP And CtrlDown = 4 And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyP And CtrlDown = 8 And cmdButton(3).Enabled
            cmdButton_Click 3
        Case vbKeyE And CtrlDown = 8 And cmdButton(4).Enabled
            cmdButton_Click 4
        Case vbKeyS And CtrlDown = 4 And cmdButton(5).Enabled
            cmdButton_Click 5
        Case vbKeyEscape
            If cmdButton(6).Enabled Then cmdButton_Click 6: Exit Function
            If cmdButton(7).Enabled Then cmdButton_Click 7
        Case vbKeyF12
            ToggleInfoPanel Me
    End Select
    
End Function

Private Sub Form_Load()

    SetUpGrid lstIconList, grdPersonsIndex
    PositionControls Me, True, grdPersonsIndex
    ColorizeControls Me, True
    ClearFields txtDescription, txtTaxNo, lblRecordCount, lblCriteria, lblSelected
    UpdateButtons Me, 7, 1, 0, 0, 0, 0, 0, 0, 1

End Sub

Private Sub grdPersonsIndex_ColHeaderMouseEnter(ByVal lCol As Long)

    grdPersonsIndex.Header.Buttons = True

End Sub

Private Sub grdPersonsIndex_ColHeaderMouseLeave(ByVal lCol As Long)

    grdPersonsIndex.Header.Buttons = False
    
End Sub

Private Sub grdPersonsIndex_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1
    
End Sub

Private Sub grdPersonsIndex_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdPersonsIndex_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeyInsert Or KeyCode = vbKeyDelete Or KeyCode = vbKeySpace Then
        grdPersonsIndex.CellIcon(grdPersonsIndex.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdPersonsIndex, KeyCode, grdPersonsIndex.CurRow, "ID"))
        lblSelected.Caption = CountSelected(grdPersonsIndex)
    End If

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdPersonsIndex", grdPersonsIndex.LayoutCol

End Sub

Private Function SaveRecords(myCurrentRow)

    Dim lngRow As Long
    Dim lngCurrentRow As Long
    Dim lngID As Long
    
    InitializeProgressBar Me, strAppTitle, grdPersonsIndex.RowCount
    
    With grdPersonsIndex
        For lngRow = 1 To .RowCount
            lngID = MainSaveRecord("CommonDB", txtTable.text, False, strAppTitle, "ID", _
                .CellValue(lngRow, "ID"), _
                .CellValue(lngRow, "Description"), _
                .CellValue(lngRow, "TaxNo"), _
                .CellValue(lngRow, "TaxOfficeID"), _
                .CellValue(lngRow, "Profession"), _
                .CellValue(lngRow, "Address"), _
                .CellValue(lngRow, "City"), _
                .CellValue(lngRow, "Phones"), _
                .CellValue(lngRow, "InCharge"), _
                .CellValue(lngRow, "VATStateID"), _
                .CellValue(lngRow, "Email"), _
                .CellValue(lngRow, "BankAccounts"), _
                .CellValue(lngRow, "CountryID"), _
                1, _
                strCurrentUser)
            If lngID = 0 Then Exit For
            UpdateProgressBar Me
        Next lngRow
    End With
    
    frmProgress.Visible = False
    
    If lngID <> 0 Then
        If MyMsgBox(1, strAppTitle, strMessages(10), 1) Then
        End If
        EnableGrid grdPersonsIndex, False, myCurrentRow
        UpdateButtons Me, 7, 0, 1, 1, 1, 1, 0, 1, 0
        blnBatchProcessing = False
    End If
   
End Function

