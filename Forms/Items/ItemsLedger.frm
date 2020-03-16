VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form ItemsLedger 
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
      Left            =   9375
      TabIndex        =   41
      Top             =   7650
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "ItemsLedger.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "ItemsLedger.frx":001C
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
         TabIndex        =   43
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
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   3615
         Index           =   0
         Left            =   150
         TabIndex        =   26
         Top             =   5100
         Width           =   9090
         Begin UserControls.newText txtCategoryShortDescription 
            Height          =   465
            Left            =   2025
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
            Left            =   2700
            TabIndex        =   27
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
            PicNormal       =   "ItemsLedger.frx":0038
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newText txtItemDescription 
            Height          =   465
            Left            =   2025
            TabIndex        =   3
            Top             =   1875
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
         Begin UserControls.newDate mskIssueFrom 
            Height          =   465
            Left            =   2025
            TabIndex        =   4
            Top             =   2400
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   820
            ForeColor       =   0
            Text            =   ""
            BackColor       =   0
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
            Left            =   3600
            TabIndex        =   5
            Top             =   2400
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   820
            ForeColor       =   0
            Text            =   ""
            BackColor       =   0
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
            Index           =   2
            Left            =   8250
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   1875
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
            PicNormal       =   "ItemsLedger.frx":05D2
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newText txtManufacturerDescription 
            Height          =   465
            Left            =   2025
            TabIndex        =   2
            Top             =   1350
            Width           =   4965
            _ExtentX        =   8758
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
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   1
            Left            =   7050
            TabIndex        =   37
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
            PicNormal       =   "ItemsLedger.frx":0B6C
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Κατασκευαστής"
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
            Index           =   4
            Left            =   450
            TabIndex        =   38
            Top             =   1425
            Width           =   1140
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
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
            Index           =   0
            Left            =   450
            TabIndex        =   35
            Top             =   2475
            Width           =   1140
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
            TabIndex        =   33
            Top             =   75
            Width           =   1665
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
            TabIndex        =   32
            Top             =   3150
            Width           =   9090
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
            Left            =   8625
            Top             =   1725
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
            Left            =   1575
            Top             =   1575
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
            Left            =   4425
            TabIndex        =   31
            Top             =   75
            Width           =   4515
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
            Left            =   3150
            TabIndex        =   30
            Top             =   900
            Width           =   4365
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
            TabIndex        =   29
            Top             =   900
            Width           =   1140
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Είδος"
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
            TabIndex        =   28
            Top             =   1950
            Width           =   1140
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
            TabIndex        =   34
            Top             =   0
            Width           =   9090
         End
      End
      Begin VB.Frame frmInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2190
         Left            =   9300
         TabIndex        =   19
         Top             =   5325
         Width           =   4515
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
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
            TabIndex        =   40
            TabStop         =   0   'False
            Text            =   "Manufacturers.ManufacturerID"
            Top             =   825
            Width           =   3540
         End
         Begin VB.TextBox txtManufacturerID 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
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
            TabIndex        =   39
            TabStop         =   0   'False
            Text            =   "3"
            Top             =   825
            Width           =   780
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
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   "4"
            Top             =   1200
            Width           =   780
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
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   "Table"
            Top             =   1200
            Width           =   3540
         End
         Begin VB.TextBox txtItemID 
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
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   "1"
            Top             =   75
            Width           =   780
         End
         Begin VB.TextBox Text3 
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
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   "Categories.CategoryID"
            Top             =   450
            Width           =   3540
         End
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
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "2"
            Top             =   450
            Width           =   780
         End
         Begin VB.TextBox Text1 
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
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "Items.ItemID"
            Top             =   75
            Width           =   3540
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   1575
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   2296
            Images          =   "ItemsLedger.frx":1106
            Version         =   131072
            KeyCount        =   2
            Keys            =   "ÿ"
         End
      End
      Begin VB.Frame frmButtonFrame 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   75
         TabIndex        =   12
         Top             =   8850
         Width           =   8940
         Begin GurhanButtonOCX.GurhanButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   13
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
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            Caption         =   "Επεξεργασία"
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
            TabIndex        =   15
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
            TabIndex        =   16
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
            TabIndex        =   17
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
            TabIndex        =   18
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
      Begin iGrid300_10Tec.iGrid grdItemsLedger 
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   75
         TabIndex        =   9
         Top             =   1125
         Width           =   2565
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Καρτέλα είδους"
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
         Width           =   3615
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
Attribute VB_Name = "ItemsLedger"
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
Dim blnPeriodIsGiven As Boolean

Dim curPrevious() As Currency
Dim curPeriod() As Currency

Private Function CalculatePreviousPeriod(myRecordset As Recordset)

    If IsDate(mskIssueFrom.text) Then
        With myRecordset
            Do While !InvoiceIssueDate < CDate(mskIssueFrom.text)
                FillArray curPrevious, _
                    CalculateDebitCreditAndBalance("Debit", "Items", !Qty, !CodeCustomers, !CodeSuppliers, !CodeInventoryQty, 0, ""), _
                    CalculateDebitCreditAndBalance("Debit", "Items", !TotalNetPostDiscount, !CodeCustomers, !CodeSuppliers, !CodeInventoryValue, 0, ""), _
                    CalculateDebitCreditAndBalance("Credit", "Items", !Qty, !CodeCustomers, !CodeSuppliers, !CodeInventoryQty, 0, ""), _
                    CalculateDebitCreditAndBalance("Credit", "Items", !TotalNetPostDiscount, !CodeCustomers, !CodeSuppliers, !CodeInventoryValue, 0, "")
                UpdateProgressBar Me
                .MoveNext
                DoEvents
                If .EOF Then
                    Exit Do
                Else
                    If Not blnProcessing Then Exit Function
                    If !InvoiceIssueDate >= CDate(mskIssueFrom.text) Then
                        Exit Do
                    End If
                End If
            Loop
            curPrevious(4) = curPrevious(0) - curPrevious(2)
            CalculatePreviousPeriod = curPrevious()
        End With
    End If

End Function

Private Function FindPersonDescription(myRefersToID, myPersonID)

    Dim rsTable As Recordset
    
    If myRefersToID = 4 Then FindPersonDescription = "": Exit Function
    
    Set rsTable = CommonDB.OpenRecordset(IIf(myRefersToID = 1, "Suppliers", "Customers"))
    With rsTable
        .Index = "ID"
        .Seek "=", myPersonID
        If Not .NoMatch Then
            FindPersonDescription = rsTable!Description
        End If
        .Close
    End With

End Function

Private Function SeekAndEditRecord(myInvoiceTrnID, myWindowTitle, myTable, myRefersTo)
    
    Dim blnFound As Boolean
    
    Select Case myRefersTo
        Case "1", "2"
            blnFound = Not SimpleSeek("Invoices", "TrnID", myInvoiceTrnID)
            If blnFound Then
                CommonTransactions.DoSharedStuff myInvoiceTrnID, myWindowTitle, myTable, myRefersTo
                If CommonTransactions.Visible Then
                    Unload Me
                    CommonTransactions.mskInvoiceIssueDate.SetFocus
                Else
                    CommonTransactions.Show 1
                End If
            Else
                DisplayMessage 17, 4, 1, ""
            End If
        Case "5"
            blnFound = Not SimpleSeek("Invoices", "TrnID", Val(myInvoiceTrnID))
            If blnFound Then
                ItemsTransactions.DoSharedStuff myInvoiceTrnID, myWindowTitle
            Else
                DisplayMessage 17, 4, 1, ""
            End If
    End Select

End Function

Private Function FindRecordsAndPopulateGrid()

    Dim blnEnableEdit As Boolean
    
    If RefreshList > 0 Then
        UpdateRecordCount lblRecordCount, lngRowCount
        UpdateCriteriaLabels txtItemDescription.text, txtManufacturerDescription.text, mskIssueFrom.text, mskIssueTo.text
        If blnPeriodIsGiven Then
            AddGridRowWithTotals grdItemsLedger, 0, "CodeDescription", strMessages(36), curPeriod(), 4, 2, 1, "QtyIn", "ValueIn", "QtyOut", "ValueOut", "QtyBalance"
            ColorizeCells grdItemsLedger, grdItemsLedger.RowCount - 1, "QtyBalance"
        End If
        AddGridRowWithTotals grdItemsLedger, 0, "CodeDescription", strMessages(32), curGrandTotal(), 4, IIf(blnPeriodIsGiven, 0, 2), 0, "QtyIn", "ValueIn", "QtyOut", "ValueOut", "QtyBalance"
        ColorizeCells grdItemsLedger, grdItemsLedger.RowCount, "QtyBalance"
        EnableGrid grdItemsLedger, False
        HighlightRow grdItemsLedger, 1, "", True
        blnEnableEdit = CheckToEnableButton(grdItemsLedger, 1, "InvoiceTrnID")
        UpdateButtons Me, 5, 0, IIf(CheckForLoadedForm("ItemsTransactions,CommonTransactions"), 0, blnEnableEdit), 1, 1, 1, 0
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
        If txtCategoryShortDescription.Enabled Then txtCategoryShortDescription.SetFocus Else mskIssueFrom.SetFocus
    End If
    
End Function

Function CreateUnicodeFile(myPrinterType, myEAFDSSString, myInvoiceHeight, myDetailLines, myTopMargin, myLeftMargin)

    On Error GoTo ErrTrap
    
    Dim lngRow As Long
    Dim intProcessedDetailLines As Integer
    
    Dim intPageNo As Integer
    
    intPageNo = 0
    intProcessedDetailLines = 0
    
    Dim curTotals(4) As Currency
    
    Open strUnicodeFile For Output As #1
    InitReport myPrinterType, myEAFDSSString, myInvoiceHeight
    GoSub Headers
    
    'Εγγραφές
    With grdItemsLedger
        For lngRow = 1 To .RowCount
            Print #1, Tab(1); .CellText(lngRow, "InvoiceIssueDate"); Tab(12); Left(.CellText(lngRow, "CodeDescription"), 30); Tab(43); .CellText(lngRow, "InvoiceNo"); Tab(50); Left(.CellText(lngRow, "PersonDescription"), 38); Tab(95 - Len(.CellText(lngRow, "QtyIn"))); .CellText(lngRow, "QtyIn"); Tab(108 - Len(.CellText(lngRow, "ValueIn"))); .CellText(lngRow, "ValueIn"); Tab(115 - Len(.CellText(lngRow, "QtyOut"))); .CellText(lngRow, "QtyOut"); Tab(128 - Len(.CellText(lngRow, "ValueOut"))); .CellText(lngRow, "ValueOut"); Tab(137 - Len(.CellText(lngRow, "QtyBalance"))); .CellText(lngRow, "QtyBalance")
            '///
            DoRunningTotal curTotals, .CellText(lngRow, "QtyIn"), .CellText(lngRow, "ValueIn"), .CellText(lngRow, "QtyOut"), .CellText(lngRow, "ValueOut"), .CellText(lngRow, "QtyIn") - .CellText(lngRow, "QtyOut")
            '///
            intProcessedDetailLines = intProcessedDetailLines + 1
            If intProcessedDetailLines > myDetailLines Then
                Print #1, ""
                AddTotalsToOutputFile Space(11) & strMessages(30), curTotals(), "095IY,108FY,115IY,128FY,137IY"
                GoSub Headers
                AddTotalsToOutputFile Space(11) & strMessages(31), curTotals(), "095IY,108FY,115IY,128FY,137IY"
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
    PrintHeadings 136, intPageNo, CustomUpperCase(lblTitle.Caption), CustomUpperCase(strCriteriaA), CustomUpperCase(strCriteriaB), myTopMargin
    PrintColumnHeadings 1, "ΗΜΕΡΟΜΗΝΙΑ", 12, "ΠΑΡΑΣΤΑΤΙΚΟ", 43, "ΝΟ", 50, "ΣΥΝΑΛΛΑΣΟΜΕΝΟΣ", 89, "---- ΕΙΣΑΓΩΓΕΣ ----", 109, "---- ΕΞΑΓΩΓΕΣ -----", 129, "ΥΠΟΛΟΙΠΟ"
    Print #1, ""
    intProcessedDetailLines = 7
    
    Return
    
ErrTrap:
    If Err.Number = 13 Then
        Resume Next
    Else
        Close #1
        CreateUnicodeFile = "Error"
        DisplayErrorMessage True, Err.Description
    End If
    
End Function

Private Function UpdateCriteriaLabels(myTireDescription, myManufacturerDescription, myIssueFrom, myIssueTo)

    strCriteriaA = "Περιγραφή " & "[ " & myTireDescription & " - " & myManufacturerDescription & " ]"
    strCriteriaB = "Εκδοση από " & IIf(myIssueFrom <> "", "[ " & myIssueFrom & " ]", "[ ΟΛΑ ]") & " έως " & IIf(myIssueTo <> "", "[ " & myIssueTo & " ]", "[ ΟΛΑ ]")
    
    lblCriteria.Caption = strCriteriaA & " " & strCriteriaB
    
End Function

Private Function UpdatePersonTable(myRefersToID)

    Select Case myRefersToID
        Case Is = 1
            UpdatePersonTable = "Suppliers"
        Case Is = 2
            UpdatePersonTable = "Customers"
    End Select

End Function

Private Function UpdateWindowTitle(myRefersToID)

    Select Case myRefersToID
        Case Is = 1
            UpdateWindowTitle = "Αγορές"
        Case Is = 2
            UpdateWindowTitle = "Πωλήσεις"
        Case Is = 3
            UpdateWindowTitle = "Κινήσεις προμηθευτών"
        Case Is = 4
            UpdateWindowTitle = "Κινήσεις πελατών"
        Case Is = 5
            UpdateWindowTitle = "Κινήσεις ειδών"
    End Select

End Function

Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
        Case 0
            If ValidateFields Then FindRecordsAndPopulateGrid
        Case 1
            With grdItemsLedger
                SeekAndEditRecord .CellText(.CurRow, "InvoiceTrnID"), .CellText(.CurRow, "WindowTitle"), .CellText(.CurRow, "PersonTable"), .CellText(.CurRow, "InvoiceRefersToID")
            End With
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
    
    'Είδος
    If DisplayMessage(1, 4, 1, "", txtItemID.text) Then txtItemDescription.SetFocus: Exit Function
    
    'Από <= Εως
    If DisplayMessage(14, 4, 1, "", mskIssueFrom.text, mskIssueTo.text) Then mskIssueFrom.SetFocus: Exit Function
    
    ValidateFields = True

End Function

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function
    
    If Not blnStatus Then
        ClearFields grdItemsLedger, lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
        frmCriteria(0).Visible = True
        If txtItemDescription.Enabled Then txtItemDescription.SetFocus Else mskIssueFrom.SetFocus
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
    Dim rstChecks As Recordset
    Dim rstItems As Recordset

    'Local μεταβλητές
    Dim lngRow As Long
    Dim lngCol As Long
    Dim blnPreviousPeriodHasBeenDisplayed As Boolean
    Dim blnAskedPeriodHasData As Boolean
    
    'Αρχικές τιμές
    ReDim curPrevious(4)
    ReDim curPeriod(4)
    ReDim curGrandTotal(4)
    
    blnPeriodIsGiven = False
    intIndex = 0
    lngRowCount = 0
    frmCriteria(0).Visible = False
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Πλέγμα
    With grdItemsLedger
        .Clear
        .Editable = False
        .Redraw = False
        .RowMode = False
    End With
    
    'Κινήσεις είδους
    strSQL = "SELECT ItemID, Qty, TotalNetPostDiscount, Invoices.InvoiceIssueDate, Invoices.InvoiceNo, Invoices.InvoiceRefersToID, Invoices.InvoiceTrnID, Invoices.InvoicePersonID, Codes.CodeDescription, Codes.CodeInventoryQty, Codes.CodeInventoryValue, Codes.CodeCustomers, Codes.CodeSuppliers " _
        & "FROM (InvoicesTrn " _
        & "INNER JOIN Invoices ON InvoicesTrn.InvoiceTrnID = Invoices.InvoiceTrnID) " _
        & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID "
    
    'Είδος
    strThisParameter = "intItem Integer"
    strThisQuery = "InvoicesTrn!ItemID = intItem"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtItemID.text)
    
    'Εκδοση
    If IsDate(mskIssueTo.text) Then
        strThisParameter = "datTo Date"
        strThisQuery = "Invoices!InvoiceIssueDate <= datTo"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskIssueTo.text)
    End If
    
    strOrder = " ORDER BY InvoiceIssueDate, InvoiceCodeID, InvoiceNo"
    
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
    
    'Θα εμφανίσω σύνολα περιόδου;
    If mskIssueFrom.text > "" Then blnPeriodIsGiven = True
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strAppTitle, rstRecordset
    
    'Προσωρινά
    UpdateButtons Me, 5, 0, 0, 0, 0, 1, 0
    cmdButton(4).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        Do While Not .EOF
            If mskIssueFrom.text <> "" And Not blnPreviousPeriodHasBeenDisplayed Then
                '///
                CalculatePreviousPeriod rstRecordset
                AddGridRowWithTotals grdItemsLedger, 0, "CodeDescription", strMessages(31), curPrevious(), 4, 1, 1, "QtyIn", "ValueIn", "QtyOut", "ValueOut", "QtyBalance"
                ColorizeCells grdItemsLedger, grdItemsLedger.RowCount - 1, "QtyBalance"
                CalculateGrandTotals curPrevious(0), curPrevious(1), curPrevious(2), curPrevious(3), curPrevious(4)
                '///
                blnPreviousPeriodHasBeenDisplayed = True
                blnAskedPeriodHasData = False
                If .EOF Then Exit Do
            End If
            grdItemsLedger.AddRow
            lngRow = grdItemsLedger.RowCount
            blnAskedPeriodHasData = True
            grdItemsLedger.CellValue(lngRow, "InvoiceTrnID") = !InvoiceTrnID
            grdItemsLedger.CellValue(lngRow, "InvoiceRefersToID") = !InvoiceRefersToID
            grdItemsLedger.CellValue(lngRow, "WindowTitle") = UpdateWindowTitle(!InvoiceRefersToID)
            grdItemsLedger.CellValue(lngRow, "PersonTable") = UpdatePersonTable(!InvoiceRefersToID)
            grdItemsLedger.CellValue(lngRow, "InvoiceIssueDate") = !InvoiceIssueDate
            grdItemsLedger.CellValue(lngRow, "CodeDescription") = !CodeDescription
            grdItemsLedger.CellValue(lngRow, "InvoiceNo") = !InvoiceNo
            grdItemsLedger.CellValue(lngRow, "PersonDescription") = FindPersonDescription(!InvoiceRefersToID, !InvoicePersonID)
            grdItemsLedger.CellValue(lngRow, "QtyIn") = CalculateDebitCreditAndBalance("Debit", "Items", !Qty, !CodeCustomers, !CodeSuppliers, !CodeInventoryQty, 0, "")
            grdItemsLedger.CellValue(lngRow, "ValueIn") = CalculateDebitCreditAndBalance("Debit", "Items", !TotalNetPostDiscount, !CodeCustomers, !CodeSuppliers, !CodeInventoryValue, 0, "")
            grdItemsLedger.CellValue(lngRow, "QtyOut") = CalculateDebitCreditAndBalance("Credit", "Items", !Qty, !CodeCustomers, !CodeSuppliers, !CodeInventoryQty, 0, "")
            grdItemsLedger.CellValue(lngRow, "ValueOut") = CalculateDebitCreditAndBalance("Credit", "Items", !TotalNetPostDiscount, !CodeCustomers, !CodeSuppliers, !CodeInventoryValue, 0, "")
            '///
            FillArray curPeriod, _
                grdItemsLedger.CellValue(lngRow, "QtyIn"), _
                grdItemsLedger.CellValue(lngRow, "ValueIn"), _
                grdItemsLedger.CellValue(lngRow, "QtyOut"), _
                grdItemsLedger.CellValue(lngRow, "ValueOut"), _
                grdItemsLedger.CellValue(lngRow, "QtyIn") - grdItemsLedger.CellValue(lngRow, "QtyOut")
            grdItemsLedger.CellValue(lngRow, "QtyBalance") = curPrevious(4) + curPeriod(4)
            ColorizeCells grdItemsLedger, lngRow, "QtyBalance"
            '///
            lngRow = lngRow + 1
            lngRowCount = lngRowCount + 1
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
        ClearFields grdItemsLedger
    Else
        '///
        CalculateGrandTotals curPeriod(0), curPeriod(1), curPeriod(2), curPeriod(3), curPeriod(4)
        RefreshList = IIf(blnAskedPeriodHasData, rstRecordset.RecordCount, 0)
        blnProcessing = False
        '///
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
    
ErrTrap:
    blnError = True
    ClearFields grdItemsLedger, frmProgress
    cmdButton(4).Caption = "Νέα αναζήτηση"
    DisplayErrorMessage True, Err.Description
        
End Function

Private Sub cmdIndex_Click(Index As Integer)

    Dim strCategoryCriteria As String
    Dim strManufacturerCriteria As String
        
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
            'Κατασκευαστής
            If txtManufacturerDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtManufacturerDescription.text, "Manufacturers", "ManufacturerDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtManufacturerID.text = tmpTableData.strCode
            txtManufacturerDescription.text = tmpTableData.strOneField
        Case 2
            'Είδος
            If txtItemDescription.text = "" Then Exit Sub
            strCategoryCriteria = IIf(txtCategoryID.text <> "", "AND CategoryID = " & txtCategoryID.text, "")
            strManufacturerCriteria = IIf(txtManufacturerID.text <> "", "AND ManufacturerID = " & txtManufacturerID.text, "")
            Set tmpRecordset = NewCheckForMatch("CommonDB", "ItemID, ItemCategoryID, ItemManufacturerID, CategoryDescription, ManufacturerDescription, ItemDescription, CategoryShortDescription", _
                "((Items", _
                "INNER JOIN Categories ON Items.ItemCategoryID = Categories.CategoryID) " & _
                "INNER JOIN Manufacturers ON Items.ItemManufacturerID = Manufacturers.ManufacturerID) ", _
                "Left(ItemQuickDescription, " & Len(txtItemDescription.text) & ") = '" & txtItemDescription.text & "'" & strCategoryCriteria & "" & strManufacturerCriteria, "", "CategoryDescription, ManufacturerDescription, ItemDescription")
            tmpTableData = DisplayIndex(tmpRecordset, True, True, "Ευρετήριο", 7, 0, 1, 2, 3, 4, 5, 6, "ID", "ID Κατηγορίας", "ID Κατασκευαστή", "Κατηγορία", "Κατασκευαστής", "Περιγραφή", "Συντ. κατηγορίας", 0, 0, 0, 40, 40, 50, 0, 1, 0, 0, 0, 0, 0, 0)
            If tmpTableData.strCode <> "" Then
                txtItemID.text = tmpTableData.strCode
                txtCategoryID.text = tmpTableData.strOneField
                txtCategoryShortDescription.text = tmpTableData.strSixField
                lblCategoryDescription.Caption = tmpTableData.strThreeField
                txtManufacturerID.text = tmpTableData.strTwoField
                txtManufacturerDescription.text = tmpTableData.strFourField
                txtItemDescription.text = tmpTableData.strFiveField
            Else
                ClearFields txtItemID, txtItemDescription
            End If
    End Select

End Sub

Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdItemsLedger, 44, GetSetting(strAppTitle, "Layout Strings", "grdItemsLedger"), _
        "05NCNInvoiceTrnID,05NCNInvoiceRefersToID,05NCNWindowTitle,05NCNPersonTable,10NCDXInvoiceIssueDate,40NLNCodeDescription,10NCNXInvoiceNo,40NLNPersonDescription,10NRIXQtyIn,10NRFXValueIn,10NRIXQtyOut,10NRFXValueOut,10NRIXQtyBalance,03NCNSelected", _
        "TrnID,RefersToID,Παράθυρο,Πίνακας συναλλασόμενων,Ημερομηνία έκδοσης,Παραστατικό,Νο παραστατικού,Συναλλασόμενος,Ποσότητα εισαγωγής,Αξία εισαγωγής,Ποσότητα εξαγωγής,Αξία εξαγωγής,Υπόλοιπο ποσότητας,Ε"
        Me.Refresh
        frmCriteria(0).Visible = True
        If txtCategoryShortDescription.Enabled Then txtCategoryShortDescription.SetFocus Else mskIssueFrom.SetFocus
    End If
    
    'AddDummyLines grdItemsLedger, 5, 5, 5, 5, 10, 40, 6, 50, 10, 10, 10, 10, 10, 3
    
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
        Case vbKeyEscape
            If cmdButton(4).Enabled Then cmdButton_Click 4: Exit Function
            If cmdButton(5).Enabled Then cmdButton_Click 5
        Case vbKeyF12
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    SetUpGrid lstIconList, grdItemsLedger
    PositionControls Me, True, grdItemsLedger
    ColorizeControls Me, True
    ClearFields txtCategoryID, txtCategoryShortDescription, lblCategoryDescription, txtItemID, txtManufacturerID, txtManufacturerDescription, txtItemDescription, mskIssueFrom, mskIssueFrom, mskIssueTo, lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
    UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1

End Sub

Private Sub grdItemsLedger_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    If grdItemsLedger.RowCount = 0 Then Exit Sub

    grdItemsLedger.RemoveRow (grdItemsLedger.RowCount): grdItemsLedger.RemoveRow (grdItemsLedger.RowCount)

End Sub

Private Sub grdItemsLedger_ColHeaderMouseEnter(ByVal lCol As Long)

    grdItemsLedger.Header.Buttons = True

End Sub

Private Sub grdItemsLedger_ColHeaderMouseLeave(ByVal lCol As Long)

    grdItemsLedger.Header.Buttons = False
    
End Sub

Private Sub grdItemsLedger_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)

    cmdButton(1).Enabled = CheckToEnableButton(grdItemsLedger, lRow, "InvoiceTrnID")
    cmdButton(1).Enabled = IIf(CheckForLoadedForm("ItemsTransactions,CommonTransactions"), 0, cmdButton(1).Enabled)

End Sub

Private Sub grdItemsLedger_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1
    
End Sub

Private Sub grdItemsLedger_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdItemsLedger_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeyInsert Or KeyCode = vbKeyDelete Or KeyCode = vbKeySpace Then
        grdItemsLedger.CellIcon(grdItemsLedger.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdItemsLedger, KeyCode, grdItemsLedger.CurRow, "InvoiceTrnID"))
        lblSelectedGridLines.Caption = CountSelected(grdItemsLedger)
        lblSelectedGridTotals.Caption = SumSelectedGridRows(grdItemsLedger, True, "QtyIn", "QtyOut", "QtyBalance")
    End If

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdItemsLedger", grdItemsLedger.LayoutCol

End Sub

Private Sub txtCategoryShortDescription_Change()

    If txtCategoryShortDescription.text = "" Then ClearFields txtCategoryID, lblCategoryDescription, txtItemID, txtItemDescription

End Sub

Private Sub txtCategoryShortDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub

Private Sub txtCategoryShortDescription_Validate(Cancel As Boolean)

    If txtCategoryID.text = "" And txtCategoryShortDescription.text <> "" Then cmdIndex_Click 0: If txtCategoryID.text = "" Then Cancel = True

End Sub

Private Sub txtItemDescription_Change()

    If txtItemDescription.text = "" Then ClearFields txtItemID

End Sub

Private Sub txtItemDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2

End Sub

Private Sub txtItemDescription_Validate(Cancel As Boolean)

    If txtItemID.text = "" And txtItemDescription.text <> "" Then cmdIndex_Click 2: If txtItemID.text = "" Then Cancel = True

End Sub

Private Sub txtManufacturerDescription_Change()

    If txtManufacturerDescription.text = "" Then ClearFields txtManufacturerID, txtItemID, txtItemDescription

End Sub

Private Sub txtManufacturerDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub

Private Sub txtManufacturerDescription_Validate(Cancel As Boolean)

    If txtManufacturerID.text = "" And txtManufacturerDescription.text <> "" Then cmdIndex_Click 1: If txtManufacturerID.text = "" Then Cancel = True

End Sub


