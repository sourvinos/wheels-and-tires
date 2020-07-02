VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form PersonsBalanceSheet 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   9765
   ClientLeft      =   0
   ClientTop       =   0
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
      Left            =   8925
      TabIndex        =   34
      Top             =   7650
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "PersonsBalanceSheet.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "PersonsBalanceSheet.frx":001C
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
         TabIndex        =   36
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
         TabIndex        =   23
         Top             =   8850
         Width           =   8940
         Begin GurhanButtonOCX.GurhanButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   24
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
            Index           =   1
            Left            =   1650
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   0
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
            MousePointer    =   99
            ShowFocusRect   =   0   'False
            BackColor       =   8438015
         End
         Begin GurhanButtonOCX.GurhanButton cmdButton 
            Height          =   690
            Index           =   5
            Left            =   7350
            TabIndex        =   26
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
            Index           =   2
            Left            =   3075
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            Caption         =   "Εκτύπωση"
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
            Index           =   4
            Left            =   5925
            TabIndex        =   28
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
         Begin GurhanButtonOCX.GurhanButton cmdButton 
            Height          =   690
            Index           =   3
            Left            =   4500
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            Caption         =   "Δημιουργία αρχείου PDF"
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
      Begin VB.Frame frmInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2190
         Left            =   8850
         TabIndex        =   12
         Tag             =   "Hidden"
         Top             =   5325
         Visible         =   0   'False
         Width           =   4515
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
            TabIndex        =   33
            TabStop         =   0   'False
            Text            =   "3"
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
            TabIndex        =   32
            TabStop         =   0   'False
            Text            =   "RefersTo"
            Top             =   825
            Width           =   1965
         End
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
            TabIndex        =   31
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
            TabIndex        =   30
            TabStop         =   0   'False
            Text            =   "2"
            Top             =   450
            Width           =   2340
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
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "Options.OptionID"
            Top             =   1200
            Width           =   1965
         End
         Begin VB.TextBox txtOptionID 
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
            Left            =   2100
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "5"
            Top             =   1200
            Width           =   2340
         End
         Begin VB.TextBox Text1 
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
            TabIndex        =   19
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
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "1"
            Top             =   75
            Width           =   2340
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   1575
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   2296
            Images          =   "PersonsBalanceSheet.frx":0038
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
         TabIndex        =   4
         Top             =   6150
         Width           =   8640
         Begin UserControls.newText txtOptionDescription 
            Height          =   465
            Left            =   1575
            TabIndex        =   3
            Top             =   1350
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   40
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
         Begin UserControls.newDate mskIssueFrom 
            Height          =   465
            Left            =   1575
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
         Begin UserControls.newDate mskIssueTo 
            Height          =   465
            Left            =   3150
            TabIndex        =   2
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
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   0
            Left            =   7800
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1350
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   820
            BackColor       =   16777215
            ButtonShape     =   3
            ButtonStyle     =   2
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            PicNormal       =   "PersonsBalanceSheet.frx":0950
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000000C0&
            Caption         =   "Εγγραφές"
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
            Index           =   0
            Left            =   450
            TabIndex        =   16
            Top             =   1425
            Width           =   690
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
            Left            =   3975
            TabIndex        =   13
            Top             =   75
            Width           =   4515
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
            Left            =   8175
            Top             =   900
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
            TabIndex        =   11
            Top             =   2100
            Width           =   8640
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
            TabIndex        =   9
            Top             =   75
            Width           =   1665
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
            TabIndex        =   5
            Top             =   900
            Width           =   690
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
            TabIndex        =   10
            Top             =   0
            Width           =   8640
         End
      End
      Begin iGrid300_10Tec.iGrid grdPersonsBalanceSheet 
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
      Begin VB.Label lblSelectedGridTotals 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Σύνολα πάνε εδώ"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   315
         Left            =   2550
         TabIndex        =   22
         Top             =   525
         Width           =   14940
      End
      Begin VB.Label lblSelectedGridLines 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
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
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2550
         TabIndex        =   15
         Top             =   825
         Width           =   14940
      End
      Begin VB.Label lblRecordCount 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Βρέθηκαν 99.999 εγγραφές"
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
         Width           =   2565
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ισοζύγιο συναλλασόμενων"
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
         Width           =   6195
      End
      Begin VB.Label lblCriteria 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
Attribute VB_Name = "PersonsBalanceSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngRowCount As Long
Dim blnError As Boolean
Dim blnProcessing As Boolean
Dim strCriteriaA As String
Dim strCriteriaB As String

Dim curPrevious() As Currency
Dim curPeriod() As Currency

Dim datLastInvoiceDate As Date

Private Function CalculatePreviousPeriod(myID, myRecordset As Recordset)

    If IsDate(mskIssueFrom.text) Then
        With myRecordset
            Do While !InvoiceIssueDate < CDate(mskIssueFrom.text) And myID = !InvoicePersonID
                GoSub StoreLastInvoiceDate
                FillArray curPrevious, _
                    CalculateDebitCreditAndBalance("Debit", "Persons", !InvoiceGrossAmount, !CodeCustomers, !CodeSuppliers, "", !PaymentWayCreditID, !ShowInList), _
                    CalculateDebitCreditAndBalance("Credit", "Persons", !InvoiceGrossAmount, !CodeCustomers, !CodeSuppliers, "", !PaymentWayCreditID, !ShowInList)
                UpdateProgressBar Me
                .MoveNext
                DoEvents
                If .EOF Then
                    Exit Do
                Else
                    If Not blnProcessing Then Exit Function
                    If !InvoiceIssueDate >= CDate(mskIssueFrom.text) Or !InvoicePersonID <> myID Then
                        Exit Do
                    End If
                End If
            Loop
            curPrevious(2) = curPrevious(0) - curPrevious(1)
            CalculatePreviousPeriod = curPrevious()
        End With
    End If
    
    Exit Function
    
StoreLastInvoiceDate:
    datLastInvoiceDate = myRecordset!InvoiceIssueDate
    
    Return

End Function

Private Function FindRecordsAndPopulateGrid()

    If RefreshList > 0 Then
        UpdateRecordCount lblRecordCount, lngRowCount
        UpdateCriteriaLabels mskIssueFrom.text, mskIssueTo.text, txtOptionDescription.text
        AddGridRowWithTotals grdPersonsBalanceSheet, 0, "PersonDescription", strMessages(32), curGrandTotal(), 5, 2, 0, "PreviousDebit", "PreviousCredit", "PreviousBalance", "Debit", "Credit", "Balance"
        ColorizeCells grdPersonsBalanceSheet, grdPersonsBalanceSheet.RowCount, "PreviousDebit", "PreviousCredit", "PreviousBalance", "Debit", "Credit", "Balance"
        EnableGrid grdPersonsBalanceSheet, False
        HighlightRow grdPersonsBalanceSheet, 1, "", True
        UpdateButtons Me, 5, 0, 1, 1, 1, 1, 0
    Else
        UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1
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
        txtOptionDescription.SetFocus
    End If
    
End Function

Function CreateUnicodeFile(myPrinterType, myEAFDSSString, myInvoiceHeight, myDetailLines, myTopMargin, myLeftMargin)

    On Error GoTo ErrTrap
    
    Dim lngRow As Long
    Dim intProcessedDetailLines As Integer
    Dim intPageNo As Integer
    
    intPageNo = 0
    intProcessedDetailLines = 0
    
    Dim curTotals(3) As Currency
    
    Open strUnicodeFile For Output As #1
    InitReport myPrinterType, myEAFDSSString, myInvoiceHeight
    GoSub Headers
    
    With grdPersonsBalanceSheet
        For lngRow = 1 To .RowCount
            Print #1, _
                Tab(1); Left(.CellText(lngRow, "PersonDescription"), 65); _
                Tab(91 - Len(.CellText(lngRow, "PreviousBalance"))); .CellText(lngRow, "PreviousBalance"); _
                Tab(106 - Len(.CellText(lngRow, "Debit"))); .CellText(lngRow, "Debit"); _
                Tab(121 - Len(.CellText(lngRow, "Credit"))); .CellText(lngRow, "Credit"); _
                Tab(136 - Len(.CellText(lngRow, "Balance"))); .CellText(lngRow, "Balance")
            '///
            DoRunningTotal curTotals, .CellText(lngRow, "PreviousBalance"), .CellText(lngRow, "Debit"), .CellText(lngRow, "Credit"), .CellText(lngRow, "Balance")
            '///
            intProcessedDetailLines = intProcessedDetailLines + 1
            If intProcessedDetailLines > Val(myDetailLines) Then
                Print #1, ""
                AddTotalsToOutputFile Space(0) & strMessages(30), curTotals(), "091FY,106FY,121FY,136FY"
                GoSub Headers
                AddTotalsToOutputFile Space(0) & strMessages(31), curTotals(), "091FY,106FY,121FY,136FY"
                Print #1, ""
                intProcessedDetailLines = intProcessedDetailLines + 2
            End If
        Next lngRow
    End With
    
    Close #1
    
    CreateUnicodeFile = strUnicodeFile
    
    Exit Function
    
Headers:
    intPageNo = intPageNo + 1
    PrintHeadings 135, intPageNo, CustomUpperCase(lblTitle.Caption), CustomUpperCase(strCriteriaA), CustomUpperCase(strCriteriaB), myTopMargin
    PrintColumnHeadings 1, "ΕΠΩΝΥΜΙΑ", 80, "ΠΡΟΗΓΟΥΜΕΝΟ", 100, "ΧΡΕΩΣΗ", 114, "ΠΙΣΤΩΣΗ", 128, "ΥΠΟΛΟΙΠΟ"
    PrintColumnHeadings 83, "ΥΠΟΛΟΙΠΟ"
    Print #1, ""
    intProcessedDetailLines = 8
    
    Return
    
ErrTrap:
    Close #1
    CreateUnicodeFile = "Error"
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function UpdateCriteriaLabels(myIssueFrom, myIssueTo, myCriteria)

    strCriteriaA = "Εκδοση από " & IIf(myIssueFrom <> "", "[ " & myIssueFrom & " ]", "[ ΟΛΑ ]") & " έως " & IIf(myIssueTo <> "", "[ " & myIssueTo & " ]", "[ ΟΛΑ ]")
    strCriteriaB = "Κριτήρια " & "[ " & myCriteria & " ]"
    
    lblCriteria.Caption = strCriteriaA & " " & strCriteriaB
    
End Function

Private Sub cmdButton_Click(Index As Integer)

    Dim strWindowTitle As String

    Select Case Index
        Case 0
            If ValidateFields Then FindRecordsAndPopulateGrid
        Case 1
            Select Case txtRefersTo.text
                Case "1"
                    strWindowTitle = "Καρτέλα προμηθευτή"
                Case "2"
                    strWindowTitle = "Καρτέλα πελάτη"
            End Select
            ShowPersonLedger _
                grdPersonsBalanceSheet.CellValue(grdPersonsBalanceSheet.CurRow, "PersonID"), _
                grdPersonsBalanceSheet.CellValue(grdPersonsBalanceSheet.CurRow, "PersonDescription"), _
                IIf(txtRefersTo.text = "1", "Καρτέλα προμηθευτή", "Καρτέλα πελάτη"), _
                txtTable.text, _
                txtOppositeTable, _
                txtRefersTo.text
        Case 2
            PrintRecords Me, "Print", False, "PrinterPrintsReportsID"
        Case 3
            PrintRecords Me, "CreatePDF", True, "PrinterPrintsReportsID"
        Case 4
            AbortProcedure False
        Case 5
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
    
    'Εγγραφές
    If DisplayMessage(1, 4, 1, "", txtOptionID.text) Then txtOptionDescription.SetFocus: Exit Function
    
    ValidateFields = True

End Function

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function
    
    If Not blnStatus Then
        ClearFields grdPersonsBalanceSheet, lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
        frmCriteria(0).Visible = True
        txtOptionDescription.SetFocus
        UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1
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

    'Local μεταβλητές
    Dim intLoop As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngPersonID As Long
    Dim strPersonDescription As String
    Dim strEmail As String
    
    'Αρχικές τιμές
    ReDim curPrevious(2)
    ReDim curPeriod(1)
    ReDim curGrandTotal(5)
    
    intIndex = 0
    lngRowCount = 0
    frmCriteria(0).Visible = False
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Πλέγμα
    With grdPersonsBalanceSheet
        .Clear
        .Editable = False
        .Redraw = False
        .RowMode = False
    End With
    
    'Αγορές, πωλήσεις, κινήσεις πελατών και προμηθευτών
    strSQL = "SELECT InvoiceID, InvoiceIssueDate, InvoiceRefersToID, InvoiceGrossAmount, InvoicePersonID, PaymentWayCreditID, CodeCustomers, CodeSuppliers, Description, Codes.ShowInList, Email " _
    & "FROM (((Invoices " _
    & "INNER JOIN " & txtTable.text & " ON Invoices.InvoicePersonID = " & txtTable.text & ".ID) " _
    & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID) " _
    & "INNER JOIN PaymentWays ON Invoices.InvoicePaymentWayID = PaymentWays.PaymentWayID) "
    
    'Αγορές = 1, Πωλήσεις = 2
    strThisParameter = "lngRefersToA Long"
    strThisQuery = "(Invoices.InvoiceRefersToID = lngRefersToA"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtRefersTo.text) - 2
    
    'Προμηθευτές = 3, Πελάτες = 4
    strThisParameter = "lngRefersToB Long"
    strThisQuery = "Invoices.InvoiceRefersToID = lngRefersToB)"
    strLogic = " OR "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtRefersTo.text)
    
    'Ποσό > μηδέν
    strThisParameter = "curGrossAmount Currency"
    strThisQuery = "Invoices.InvoiceGrossAmount > curGrossAmount"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = 0
    
    'Συναλλασόμενος
    strThisParameter = "lngPersonID Long"
    strThisQuery = "Invoices.InvoicePersonID <> lngPersonID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = 0
    
    'Εως
    If IsDate(mskIssueTo.text) Then
        strThisParameter = "datIssueTo Date"
        strThisQuery = "Invoices.InvoiceIssueDate <= datIssueTo"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskIssueTo.text)
    End If
        
    'Ταξινόμηση
    strOrder = " ORDER BY InvoicePersonID, InvoiceIssueDate"
    
    'Προσθέτω τα κριτήρια
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
    End If
    
    'SQL
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
    UpdateButtons Me, 5, 0, 0, 0, 0, 1, 0
    cmdButton(4).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    '1η εγγραφή
    GoSub UpdateAreas
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        Do While Not .EOF
            If !InvoicePersonID = lngPersonID Then
                GoSub StoreLastInvoiceDate
                If !InvoiceIssueDate < CDate(mskIssueFrom.text) Then
                    CalculatePreviousPeriod lngPersonID, rstRecordset
                    If Not blnProcessing Then Exit Do
                Else
                    FillArray curPeriod, _
                        CalculateDebitCreditAndBalance("Debit", "Persons", !InvoiceGrossAmount, !CodeCustomers, !CodeSuppliers, "", !PaymentWayCreditID, !ShowInList), _
                        CalculateDebitCreditAndBalance("Credit", "Persons", !InvoiceGrossAmount, !CodeCustomers, !CodeSuppliers, "", !PaymentWayCreditID, !ShowInList)
                    UpdateProgressBar Me
                    rstRecordset.MoveNext
                    DoEvents
                    If Not blnProcessing Then Exit Do
                End If
            Else
                If txtOptionID.text = "1" Or (txtOptionID.text = "2" And (curPrevious(2) + curPeriod(0) - curPeriod(1) <> 0)) Or (txtOptionID.text = "3" And (curPeriod(0) <> 0 Or curPeriod(1) <> 0)) Then
                    GoSub AddLine
                    ColorizeCells grdPersonsBalanceSheet, lngRow, "PreviousBalance", "Debit", "Credit", "Balance"
                    CalculateGrandTotals curPrevious(0), curPrevious(1), curPrevious(2), curPeriod(0), curPeriod(1), curPrevious(2) + curPeriod(0) - curPeriod(1)
                End If
                ClearVariables curPrevious(0), curPrevious(1), curPrevious(2), curPeriod(0), curPeriod(1)
                GoSub UpdateAreas
            End If
        Loop
        If blnProcessing Then
            If .RecordCount <> 0 Then
                .MoveLast
                If txtOptionID.text = "1" Or (txtOptionID.text = "2" And (curPrevious(2) + curPeriod(0) - curPeriod(1) <> 0)) Or (txtOptionID.text = "3" And (curPeriod(0) <> 0 Or curPeriod(1) <> 0)) Then
                GoSub StoreLastInvoiceDate
                    GoSub AddLine
                    ColorizeCells grdPersonsBalanceSheet, lngRow, "PreviousBalance", "Debit", "Credit", "Balance"
                    CalculateGrandTotals curPrevious(0), curPrevious(1), curPrevious(2), curPeriod(0), curPeriod(1), curPrevious(2) + curPeriod(0) - curPeriod(1)
                End If
            End If
        End If
    End With
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        RefreshList = 0
        ClearFields grdPersonsBalanceSheet
    Else
        grdPersonsBalanceSheet.Sort ("PersonDescription")
        RefreshList = lngRowCount
        blnProcessing = False
    End If
    
    'Τελικές ενέργειες
    cmdButton(4).Caption = "Νέα αναζήτηση"
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
    
AddLine:
    With grdPersonsBalanceSheet
        .AddRow
        lngRow = .RowCount
        .CellValue(.RowCount, "PersonID") = lngPersonID
        .CellValue(.RowCount, "PersonDescription") = strPersonDescription
        .CellValue(.RowCount, "Email") = strEmail
        .CellValue(.RowCount, "LastInvoiceIssueDate") = datLastInvoiceDate
        .CellValue(.RowCount, "PreviousDebit") = curPrevious(0)
        .CellValue(.RowCount, "PreviousCredit") = curPrevious(1)
        .CellValue(.RowCount, "PreviousBalance") = curPrevious(2)
        .CellValue(.RowCount, "Debit") = curPeriod(0)
        .CellValue(.RowCount, "Credit") = curPeriod(1)
        .CellValue(.RowCount, "Balance") = curPrevious(2) + curPeriod(0) - curPeriod(1)
        lngRowCount = lngRowCount + 1
    End With
    
    Return

UpdateAreas:
    lngPersonID = rstRecordset!InvoicePersonID
    strPersonDescription = rstRecordset!Description
    datLastInvoiceDate = rstRecordset!InvoiceIssueDate
    strEmail = rstRecordset!Email
    
    Return
    
StoreLastInvoiceDate:
    datLastInvoiceDate = rstRecordset!InvoiceIssueDate
    
    Return
    
ErrTrap:
    blnError = True
    ClearFields grdPersonsBalanceSheet, frmProgress
    cmdButton(4).Caption = "Νέα αναζήτηση"
    DisplayErrorMessage True, Err.Description
        
End Function

Private Sub cmdIndex_Click(Index As Integer)

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case Index
        Case 0
            'Εγγραφές
            If txtOptionDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtOptionDescription.text, "Options", "OptionDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtOptionID.text = tmpTableData.strCode
            txtOptionDescription.text = tmpTableData.strOneField
    End Select

End Sub

Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdPersonsBalanceSheet, 44, GetSetting(strAppTitle, "Layout Strings", "grdPersonsBalanceSheet"), _
            "10NCNPersonID,50NLNPersonDescription,50NLNEmail,10NCDXLastInvoiceIssueDate,10NRFXPreviousDebit,10NRFXPreviousCredit,10NRFXPreviousBalance,10NRFDebit,10NRFCredit,10NRFBalance,03NCNSelected", _
            "ID,Επωνυμία,Email,Τελευταία εγγραφή,Προηγούμενη χρέωση,Προηγούμενη πίστωση,Προηγούμενο υπόλοιπο,Χρέωση,Πίστωση,Υπόλοιπο,Ε"
        Me.Refresh
        frmCriteria(0).Visible = True
        mskIssueFrom.SetFocus
    End If
    
    'AddDummyLines grdPersonsBalanceSheet, 6, 50, 30, 10, 13, 13, 13, 13, 13, 13, 4
    
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
        Case vbKeyL And CtrlDown = 4 And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyP And CtrlDown = 4 And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyP And CtrlDown = 8 And cmdButton(3).Enabled
            cmdButton_Click 3
        Case vbKeyEscape
            If cmdButton(4).Enabled Then cmdButton_Click 4: Exit Function
            If cmdButton(5).Enabled Then cmdButton_Click 5
        Case vbKeyF12
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    SetUpGrid lstIconList, grdPersonsBalanceSheet
    PositionControls Me, True, grdPersonsBalanceSheet
    ColorizeControls Me, True
    ClearFields mskIssueFrom, mskIssueTo, txtOptionID, txtOptionDescription, lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
    UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1
    
End Sub

Private Sub grdPersonsBalanceSheet_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    If grdPersonsBalanceSheet.RowCount = 0 Then Exit Sub

    grdPersonsBalanceSheet.RemoveRow (grdPersonsBalanceSheet.RowCount): grdPersonsBalanceSheet.RemoveRow (grdPersonsBalanceSheet.RowCount)

End Sub

Private Sub grdPersonsBalanceSheet_ColHeaderMouseEnter(ByVal lCol As Long)

    grdPersonsBalanceSheet.Header.Buttons = True

End Sub

Private Sub grdPersonsBalanceSheet_ColHeaderMouseLeave(ByVal lCol As Long)

    grdPersonsBalanceSheet.Header.Buttons = False
    
End Sub

Private Sub grdPersonsBalanceSheet_ContentsSorted()

    AddGridRowWithTotals grdPersonsBalanceSheet, 0, "PersonDescription", strMessages(32), curGrandTotal(), 5, 2, 0, "PreviousDebit", "PreviousCredit", "PreviousBalance", "Debit", "Credit", "Balance"

End Sub

Private Sub grdPersonsBalanceSheet_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)

    cmdButton(1).Enabled = CheckToEnableButton(grdPersonsBalanceSheet, lRow, "PersonID")

End Sub

Private Sub grdPersonsBalanceSheet_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1

End Sub


Private Sub grdPersonsBalanceSheet_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdPersonsBalanceSheet_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeyInsert Or KeyCode = vbKeyDelete Or KeyCode = vbKeySpace Then
        grdPersonsBalanceSheet.CellIcon(grdPersonsBalanceSheet.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdPersonsBalanceSheet, KeyCode, grdPersonsBalanceSheet.CurRow, "PersonID"))
        lblSelectedGridLines.Caption = CountSelected(grdPersonsBalanceSheet)
        lblSelectedGridTotals.Caption = SumSelectedGridRows(grdPersonsBalanceSheet, False, "PreviousBalance", "Debit", "Credit", "Balance")
    End If

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdPersonsBalanceSheet", grdPersonsBalanceSheet.LayoutCol

End Sub

Private Sub txtOptionDescription_Change()

    If txtOptionDescription.text = "" Then ClearFields txtOptionID
    
End Sub

Private Sub txtOptionDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub

Private Sub txtOptionDescription_Validate(Cancel As Boolean)

    If txtOptionID.text = "" And txtOptionDescription.text <> "" Then cmdIndex_Click 0: If txtOptionID.text = "" Then Cancel = True

End Sub

