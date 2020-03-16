VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form itemsBalanceSheet 
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
      Left            =   9075
      TabIndex        =   31
      Top             =   7650
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "itemsBalanceSheet.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "itemsBalanceSheet.frx":001C
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
         TabIndex        =   33
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
         TabIndex        =   24
         Top             =   8850
         Width           =   8940
         Begin GurhanButtonOCX.GurhanButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   25
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
            TabIndex        =   26
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
            TabIndex        =   27
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
            TabIndex        =   28
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
            TabIndex        =   29
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
            TabIndex        =   30
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
         Height          =   1815
         Left            =   9000
         TabIndex        =   13
         Tag             =   "Hidden"
         Top             =   5700
         Visible         =   0   'False
         Width           =   4515
         Begin VB.TextBox txtCategoryID 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   38
            TabStop         =   0   'False
            Text            =   "2"
            Top             =   450
            Width           =   780
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   37
            TabStop         =   0   'False
            Text            =   "Categories.CategoryID"
            Top             =   450
            Width           =   3540
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
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   "Options.OptionID"
            Top             =   825
            Width           =   3540
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
            Left            =   3675
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "4"
            Top             =   825
            Width           =   780
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
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "Table"
            Top             =   75
            Width           =   3540
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
            Left            =   3675
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "1"
            Top             =   75
            Width           =   780
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   1200
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   2296
            Images          =   "itemsBalanceSheet.frx":0038
            Version         =   131072
            KeyCount        =   2
            Keys            =   "ÿ"
         End
      End
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   3090
         Index           =   0
         Left            =   225
         TabIndex        =   5
         Top             =   5625
         Width           =   8715
         Begin UserControls.newText txtOptionDescription 
            Height          =   465
            Left            =   1650
            TabIndex        =   2
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
            Left            =   1650
            TabIndex        =   3
            Top             =   1875
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
            TabIndex        =   4
            Top             =   1875
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
            Index           =   1
            Left            =   7875
            TabIndex        =   18
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
            PicNormal       =   "itemsBalanceSheet.frx":0950
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newText txtCategoryShortDescription 
            Height          =   465
            Left            =   1650
            TabIndex        =   1
            Top             =   825
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   820
            Alignment       =   2
            ForeColor       =   0
            MaxLength       =   2
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
            Left            =   2325
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   825
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
            PicNormal       =   "itemsBalanceSheet.frx":0EEA
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Κατηγορία"
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
            Index           =   7
            Left            =   450
            TabIndex        =   36
            Top             =   900
            Width           =   765
         End
         Begin VB.Label lblCategoryDescription 
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ"
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
            Left            =   2775
            TabIndex        =   35
            Top             =   900
            Width           =   4365
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
            TabIndex        =   17
            Top             =   1425
            Width           =   765
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
            Left            =   4050
            TabIndex        =   14
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
            Left            =   1200
            Top             =   1275
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
            Left            =   8250
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
            TabIndex        =   12
            Top             =   2625
            Width           =   8715
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
            TabIndex        =   10
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
            TabIndex        =   6
            Top             =   1950
            Width           =   765
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
            TabIndex        =   11
            Top             =   0
            Width           =   8715
         End
      End
      Begin iGrid300_10Tec.iGrid grditemsBalanceSheet 
         Height          =   7290
         Left            =   75
         TabIndex        =   7
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
         TabIndex        =   23
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   1125
         Width           =   2565
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ισοζύγιο ειδών"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   30
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   720
         Left            =   75
         TabIndex        =   9
         Top             =   75
         Width           =   3315
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
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   2550
         TabIndex        =   8
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
Attribute VB_Name = "itemsBalanceSheet"
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

Private Function CalculatePreviousPeriod(myID, myRecordset As Recordset)

    If IsDate(mskIssueFrom.text) Then
        With myRecordset
            Do While !InvoiceIssueDate < CDate(mskIssueFrom.text) And myID = !ItemID
                FillArray curPrevious, _
                    CalculateDebitCreditAndBalance("Debit", "Items", !Qty, "", "", !CodeInventoryQty, "", ""), _
                    CalculateDebitCreditAndBalance("Credit", "Items", !Qty, "", "", !CodeInventoryQty, "", "")
                UpdateProgressBar Me
                .MoveNext
                DoEvents
                If .EOF Then
                    Exit Do
                Else
                    If Not blnProcessing Then Exit Function
                    If !InvoiceIssueDate >= CDate(mskIssueFrom.text) Or !ItemID <> myID Then
                        Exit Do
                    End If
                End If
            Loop
            curPrevious(2) = curPrevious(0) - curPrevious(1)
            CalculatePreviousPeriod = curPrevious()
        End With
    End If

End Function

Private Function FindRecordsAndPopulateGrid()

    If RefreshList > 0 Then
        UpdateRecordCount lblRecordCount, lngRowCount
        UpdateCriteriaLabels lblCategoryDescription, mskIssueFrom.text, mskIssueTo.text, txtOptionDescription.text
        AddGridRowWithTotals grditemsBalanceSheet, 0, "ItemDescription", strMessages(32), curGrandTotal(), 7, 2, 0, "PreviousQtyIn", "PreviousQtyOut", "PreviousQtyBalance", "QtyIn", "ValueIn", "QtyOut", "ValueOut", "QtyBalance"
        ColorizeCells grditemsBalanceSheet, grditemsBalanceSheet.RowCount, "PreviousQtyIn", "PreviousQtyOut", "PreviousQtyBalance", "QtyIn", "ValueIn", "QtyOut", "ValueOut", "QtyBalance"
        EnableGrid grditemsBalanceSheet, False
        HighlightRow grditemsBalanceSheet, 1, "", True
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
        txtCategoryShortDescription.SetFocus
    End If
    
End Function

Function CreateUnicodeFile(myPrinterType, myEAFDSSString, myInvoiceHeight, myDetailLines, myTopMargin, myLeftMargin)

    On Error GoTo ErrTrap
    
    Dim lngRow As Long
    Dim intProcessedDetailLines As Integer
    Dim intPageNo As Integer
    
    intPageNo = 0
    intProcessedDetailLines = 0
    
    Dim curTotals(5) As Currency
    
    Open strUnicodeFile For Output As #1
    InitReport myPrinterType, myEAFDSSString, myInvoiceHeight
    GoSub Headers
    
    With grditemsBalanceSheet
        For lngRow = 1 To .RowCount
            Print #1, _
                Tab(7 - Len(.CellText(lngRow, "ItemID"))); .CellText(lngRow, "ItemID"); _
                Tab(8); Left(.CellText(lngRow, "ItemDescription"), 40); _
                Tab(49); Left(.CellText(lngRow, "ManufacturerDescription"), 25); _
                Tab(83 - Len(.CellText(lngRow, "PreviousQtyBalance"))); .CellText(lngRow, "PreviousQtyBalance"); _
                Tab(91 - Len(.CellText(lngRow, "QtyIn"))); .CellText(lngRow, "QtyIn"); _
                Tab(105 - Len(.CellText(lngRow, "ValueIn"))); .CellText(lngRow, "ValueIn"); _
                Tab(113 - Len(.CellText(lngRow, "QtyOut"))); .CellText(lngRow, "QtyOut"); _
                Tab(126 - Len(.CellText(lngRow, "ValueOut"))); .CellText(lngRow, "ValueOut"); _
                Tab(136 - Len(.CellText(lngRow, "QtyBalance"))); .CellText(lngRow, "QtyBalance")
            '///
            DoRunningTotal curTotals, .CellText(lngRow, "PreviousQtyBalance"), .CellText(lngRow, "QtyIn"), .CellText(lngRow, "ValueIn"), .CellText(lngRow, "QtyOut"), .CellText(lngRow, "ValueOut"), .CellText(lngRow, "QtyBalance")
            '///
            intProcessedDetailLines = intProcessedDetailLines + 1
            If intProcessedDetailLines > Val(myDetailLines) Then
                Print #1, ""
                AddTotalsToOutputFile Space(7) & strMessages(30), curTotals(), "083IY,091IY,105FY,113IY,126FY,136IY"
                GoSub Headers
                AddTotalsToOutputFile Space(7) & strMessages(31), curTotals(), "083IY,091IY,105FY,113IY,126FY,136IY"
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
    PrintColumnHeadings 5, "ID", 8, "ΠΕΡΙΓΡΑΦΗ", 49, "ΚΑΤΑΣΚΕΥΑΣΤΗΣ", 75, "ΠΡΟΗΓ/ΝΟ", 85, "----- ΕΙΣΑΓΩΓΕΣ ----", 107, "----- ΕΞΑΓΩΓΕΣ ----", 128, "ΥΠΟΛΟΙΠΟ"
    PrintColumnHeadings 75, "ΥΠΟΛΟΙΠΟ", 85, "ΠΟΣ/ΤΑ", 101, "ΑΞΙΑ", 107, "ΠΟΣ/ΤΑ", 122, "ΑΞΙΑ", 129, "ΠΟΣ/ΤΑΣ"
    Print #1, ""
    intProcessedDetailLines = 8
    
    Return
    
ErrTrap:
    Close #1
    CreateUnicodeFile = "Error"
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function ShowLedger(myRow)

    With ItemsLedger
        .txtCategoryID.text = txtCategoryID.text
        .txtCategoryShortDescription.text = txtCategoryShortDescription.text
        .lblCategoryDescription.Caption = lblCategoryDescription.Caption
        .txtManufacturerID.text = grditemsBalanceSheet.CellValue(myRow, "ManufacturerID")
        .txtManufacturerDescription.text = grditemsBalanceSheet.CellValue(myRow, "ManufacturerDescription")
        .txtItemID.text = grditemsBalanceSheet.CellText(grditemsBalanceSheet.CurRow, "ItemID")
        .txtItemDescription.text = grditemsBalanceSheet.CellText(grditemsBalanceSheet.CurRow, "ItemDescription")
        .txtTable.text = txtTable.text
        .Tag = "True"
        DisableFields .txtCategoryShortDescription, .txtManufacturerDescription, .txtItemDescription, .cmdIndex(0), .cmdIndex(1), .cmdIndex(2)
        .Show 1, Me
    End With
    
End Function

Private Function UpdateCriteriaLabels(myCategory, myIssueFrom, myIssueTo, myCriteria)

    strCriteriaA = "Κατηγορία [ " & myCategory & " ] Εκδοση από " & IIf(myIssueFrom <> "", "[ " & myIssueFrom & " ]", "[ ΟΛΑ ]") & " έως " & IIf(myIssueTo <> "", "[ " & myIssueTo & " ]", "[ ΟΛΑ ]")
    strCriteriaB = "Κριτήρια " & "[ " & myCriteria & " ]"
    
    lblCriteria.Caption = strCriteriaA & " " & strCriteriaB
    
End Function


Private Function UpdateWindowTitle(myRefersToID)

    Select Case myRefersToID
        Case Is = 3
            UpdateWindowTitle = "Καρτέλα προμηθευτή"
        Case Is = 4
            UpdateWindowTitle = "Καρτέλα πελάτη"
    End Select

End Function

Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
        Case 0
            If ValidateFields Then FindRecordsAndPopulateGrid
        Case 1
            ShowLedger grditemsBalanceSheet.CurRow
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
    
    'Κατηγορία
    If DisplayMessage(1, 4, 1, "", txtCategoryID.text) Then txtCategoryShortDescription.SetFocus: Exit Function
    
    'Εγγραφές
    If DisplayMessage(1, 4, 1, "", txtOptionID.text) Then txtOptionDescription.SetFocus: Exit Function
    
    'Από
    If DisplayMessage(1, 4, 1, "", mskIssueFrom.text) Then mskIssueFrom.SetFocus: Exit Function
    
    'Εως
    If DisplayMessage(1, 4, 1, "", mskIssueTo.text) Then mskIssueTo.SetFocus: Exit Function
    
    'Εκδοση
    If DisplayMessage(14, 4, 1, "", mskIssueFrom.text, mskIssueTo.text) Then mskIssueFrom.SetFocus: Exit Function
    
    ValidateFields = True

End Function


Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function
    
    If Not blnStatus Then
        ClearFields grditemsBalanceSheet, lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
        frmCriteria(0).Visible = True
        txtCategoryShortDescription.SetFocus
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
    Dim lngItemID As Long
    Dim strItemDescription As String
    Dim lngManufacturerID As Long
    Dim strManufacturerDescription As String
    
    'Αρχικές τιμές
    ReDim curPrevious(2)
    ReDim curPeriod(3)
    ReDim curGrandTotal(7)
    
    intIndex = 0
    lngRowCount = 0
    frmCriteria(0).Visible = False
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Πλέγμα
    With grditemsBalanceSheet
        .Clear
        .Editable = False
        .Redraw = False
        .RowMode = False
    End With
    
    'Αγορές, πωλήσεις, κινήσεις πελατών και προμηθευτών
    strSQL = "SELECT InvoiceIssueDate, InvoicesTrn.ItemID, Qty, TotalNetPostDiscount, InvoicesTrn.InvoiceTrnID, ItemDescription, ManufacturerID, ManufacturerDescription, CodeInventoryQty, CodeInventoryValue " _
    & "FROM (((InvoicesTrn " _
    & "INNER JOIN Invoices ON InvoicesTrn.InvoiceTrnID = Invoices.InvoiceTrnID) " _
    & "INNER JOIN Items ON InvoicesTrn.ItemID = Items.ItemID) " _
    & "INNER JOIN Manufacturers ON Items.ItemManufacturerID = Manufacturers.ManufacturerID) " _
    & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID "
    
    'Κατηγορία
    strThisParameter = "lngCategoryID Long"
    strThisQuery = "Items.ItemCategoryID = lngCategoryID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtCategoryID.text)
    
    'Είδος
    'strThisParameter = "lngItemID Long"
    'strThisQuery = "Items.ItemID = lngItemID"
    'strLogic = " AND "
    'GoSub UpdateSQLString
    'arrQuery(intIndex) = "302"
    
    'Εως
    If IsDate(mskIssueTo.text) Then
        strThisParameter = "datIssueTo Date"
        strThisQuery = "Invoices.InvoiceIssueDate <= datIssueTo"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskIssueTo.text)
    End If
        
    'Ταξινόμηση
    strOrder = " ORDER BY InvoicesTrn.ItemID, InvoiceIssueDate"
    
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
            If !ItemID = lngItemID Then
                If !InvoiceIssueDate < CDate(mskIssueFrom.text) Then
                    CalculatePreviousPeriod lngItemID, rstRecordset
                    If Not blnProcessing Then Exit Do
                Else
                    FillArray curPeriod, _
                        CalculateDebitCreditAndBalance("Debit", "Items", !Qty, "", "", !CodeInventoryQty, "", ""), _
                        CalculateDebitCreditAndBalance("Debit", "Items", !TotalNetPostDiscount, "", "", !CodeInventoryValue, "", ""), _
                        CalculateDebitCreditAndBalance("Credit", "Items", !Qty, "", "", !CodeInventoryQty, "", ""), _
                        CalculateDebitCreditAndBalance("Credit", "Items", !TotalNetPostDiscount, "", "", !CodeInventoryValue, "", "")
                    UpdateProgressBar Me
                    rstRecordset.MoveNext
                    DoEvents
                    If Not blnProcessing Then Exit Do
                End If
            Else
                If txtOptionID.text = "1" Or (txtOptionID.text = "2" And curPeriod(3) <> 0) Or (txtOptionID.text = "3" And (curPeriod(0) <> 0 Or curPeriod(1) <> 0)) Then
                    GoSub AddLine
                    ColorizeCells grditemsBalanceSheet, lngRow, "PreviousQtyIn", "PreviousQtyOut", "PreviousQtyBalance", "QtyIn", "ValueIn", "QtyOut", "ValueOut", "QtyBalance"
                    CalculateGrandTotals curPrevious(0), curPrevious(1), curPrevious(2), curPeriod(0), curPeriod(1), curPeriod(2), curPeriod(3), curPrevious(2) + curPeriod(0) - curPeriod(2)
                End If
                ClearVariables curPrevious(0), curPrevious(1), curPrevious(2), curPeriod(0), curPeriod(1), curPeriod(2), curPeriod(3)
                GoSub UpdateAreas
            End If
        Loop
        If blnProcessing Then
            If txtOptionID.text = "1" Or (txtOptionID.text = "2" And curPeriod(3) <> 0) Or (txtOptionID.text = "3" And (curPeriod(0) <> 0 Or curPeriod(1) <> 0)) Then
                GoSub AddLine
                ColorizeCells grditemsBalanceSheet, lngRow, "PreviousQtyIn", "PreviousQtyOut", "PreviousQtyBalance", "QtyIn", "ValueIn", "QtyOut", "ValueOut", "QtyBalance"
                CalculateGrandTotals curPrevious(0), curPrevious(1), curPrevious(2), curPeriod(0), curPeriod(1), curPeriod(2), curPeriod(3), curPrevious(2) + curPeriod(0) - curPeriod(2)
            End If
        End If
    End With
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        RefreshList = 0
        ClearFields grditemsBalanceSheet
    Else
        grditemsBalanceSheet.Sort Array("ManufacturerDescription", "ItemDescription")
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
    With grditemsBalanceSheet
        .AddRow
        lngRow = .RowCount
        .CellValue(.RowCount, "ItemID") = lngItemID
        .CellValue(.RowCount, "ItemDescription") = strItemDescription
        .CellValue(.RowCount, "ManufacturerID") = lngManufacturerID
        .CellValue(.RowCount, "ManufacturerDescription") = strManufacturerDescription
        .CellValue(.RowCount, "PreviousQtyIn") = curPrevious(0)
        .CellValue(.RowCount, "PreviousQtyOut") = curPrevious(1)
        .CellValue(.RowCount, "PreviousQtyBalance") = curPrevious(2)
        .CellValue(.RowCount, "QtyIn") = curPeriod(0)
        .CellValue(.RowCount, "ValueIn") = curPeriod(1)
        .CellValue(.RowCount, "QtyOut") = curPeriod(2)
        .CellValue(.RowCount, "ValueOut") = curPeriod(3)
        .CellValue(.RowCount, "QtyBalance") = curPrevious(2) + curPeriod(0) - curPeriod(2)
        lngRowCount = lngRowCount + 1
    End With
    
    Return

UpdateAreas:
    lngItemID = rstRecordset!ItemID
    strItemDescription = rstRecordset!ItemDescription
    lngManufacturerID = rstRecordset!ManufacturerID
    strManufacturerDescription = rstRecordset!ManufacturerDescription
    
    Return

ErrTrap:
    blnError = True
    ClearFields grditemsBalanceSheet, frmProgress
    cmdButton(4).Caption = "Νέα αναζήτηση"
    DisplayErrorMessage True, Err.Description
        
End Function

Private Sub cmdIndex_Click(Index As Integer)

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case Index
        Case 0
            'Κατηγορία
            If txtCategoryShortDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtCategoryShortDescription.text, "Categories", "CategoryShortDescription", "String", 1, 3)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 3, 0, 1, 2, "ID", "Συντ.", "Περιγραφή", 0, 4, 40, 1, 1, 0)
            txtCategoryID.text = tmpTableData.strCode
            txtCategoryShortDescription.text = tmpTableData.strOneField
            lblCategoryDescription.Caption = tmpTableData.strTwoField
        Case 1
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
        AddColumnsToGrid grditemsBalanceSheet, 44, GetSetting(strAppTitle, "Layout Strings", "grditemsBalanceSheet"), _
            "06NCNItemID,50NLNItemDescription,06NCNManufacturerID,40NLNManufacturerDescription,10NRIXPreviousQtyIn,10NRIXPreviousQtyOut,10NRIXPreviousQtyBalance,10NRIXQtyIn,10NRFXValueIn,10NRIXQtyOut,10NRFXValueOut,10NRIXQtyBalance,03NCNSelected", _
            "ID,Περιγραφή,ID Κατασκευαστή,Κατασκευαστής,Προηγούμενη εισαγωγή,Προηγούμενη εξαγωγή,Προηγούμενο υπόλοιπο,Ποσότητα εισαγωγής,Αξία εισαγωγής,Ποσότητα εξαγωγής,Αξία εξαγωγής,Υπόλοιπο ποσότητας,Ε"
        Me.Refresh
        frmCriteria(0).Visible = True
        txtCategoryShortDescription.SetFocus
    End If
    
    'AddDummyLines grditemsBalanceSheet, 6, 50, 6, 40, 10, 10, 10, 10, 10, 10, 10, 10, 3
    
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

    SetUpGrid lstIconList, grditemsBalanceSheet
    PositionControls Me, True, grditemsBalanceSheet
    ColorizeControls Me, True
    ClearFields txtCategoryID, txtCategoryShortDescription, lblCategoryDescription, txtOptionID, txtOptionDescription, mskIssueFrom, mskIssueTo, lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
    UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1

End Sub

Private Sub grditemsBalanceSheet_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    If grditemsBalanceSheet.RowCount = 0 Then Exit Sub

    grditemsBalanceSheet.RemoveRow (grditemsBalanceSheet.RowCount): grditemsBalanceSheet.RemoveRow (grditemsBalanceSheet.RowCount)

End Sub

Private Sub grditemsBalanceSheet_ColHeaderMouseEnter(ByVal lCol As Long)

    grditemsBalanceSheet.Header.Buttons = True

End Sub

Private Sub grditemsBalanceSheet_ColHeaderMouseLeave(ByVal lCol As Long)

    grditemsBalanceSheet.Header.Buttons = False
    
End Sub

Private Sub grditemsBalanceSheet_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)

    cmdButton(1).Enabled = CheckToEnableButton(grditemsBalanceSheet, lRow, "ItemID")

End Sub

Private Sub grditemsBalanceSheet_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1

End Sub


Private Sub grditemsBalanceSheet_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grditemsBalanceSheet_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeyInsert Or KeyCode = vbKeyDelete Or KeyCode = vbKeySpace Then
        grditemsBalanceSheet.CellIcon(grditemsBalanceSheet.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grditemsBalanceSheet, KeyCode, grditemsBalanceSheet.CurRow, "ItemID"))
        lblSelectedGridLines.Caption = CountSelected(grditemsBalanceSheet)
        lblSelectedGridTotals.Caption = SumSelectedGridRows(grditemsBalanceSheet, False, "PreviousQtyBalance", "QtyIn", "ValueIn", "QtyOut", "ValueOut", "QtyBalance")
    End If

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grditemsBalanceSheet", grditemsBalanceSheet.LayoutCol

End Sub

Private Sub txtCategoryShortDescription_Change()

    If txtCategoryShortDescription.text = "" Then ClearFields txtCategoryID, lblCategoryDescription

End Sub

Private Sub txtCategoryShortDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub


Private Sub txtCategoryShortDescription_Validate(Cancel As Boolean)

    If txtCategoryID.text = "" And txtCategoryShortDescription.text <> "" Then cmdIndex_Click 0: If txtCategoryID.text = "" Then Cancel = True

End Sub


Private Sub txtOptionDescription_Change()

    If txtOptionDescription.text = "" Then ClearFields txtOptionID
    
End Sub

Private Sub txtOptionDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1

End Sub

Private Sub txtOptionDescription_Validate(Cancel As Boolean)

    If txtOptionID.text = "" And txtOptionDescription.text <> "" Then cmdIndex_Click 1: If txtOptionID.text = "" Then Cancel = True

End Sub

