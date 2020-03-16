VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form ItemsIndex 
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
         Picture         =   "ItemsIndex.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "ItemsIndex.frx":001C
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
         TabIndex        =   17
         Top             =   8850
         Width           =   8940
         Begin GurhanButtonOCX.GurhanButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   18
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
            TabIndex        =   19
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
            Index           =   5
            Left            =   7350
            TabIndex        =   20
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
            TabIndex        =   21
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
            TabIndex        =   22
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
            TabIndex        =   23
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
         Left            =   8850
         TabIndex        =   11
         Top             =   5700
         Width           =   4515
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
            TabIndex        =   27
            TabStop         =   0   'False
            Text            =   "Items.ItemManufacturerID"
            Top             =   450
            Width           =   3540
         End
         Begin VB.TextBox txtCategoryID 
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
            TabIndex        =   26
            TabStop         =   0   'False
            Text            =   "1"
            Top             =   75
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
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   "Items.CategoryID"
            Top             =   75
            Width           =   3540
         End
         Begin VB.TextBox txtManufacturerID 
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
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   "2"
            Top             =   450
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
            TabIndex        =   14
            TabStop         =   0   'False
            Text            =   "Table"
            Top             =   825
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
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "3"
            Top             =   825
            Width           =   780
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   1200
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   2296
            Images          =   "ItemsIndex.frx":0038
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
         Left            =   150
         TabIndex        =   4
         Top             =   5625
         Width           =   8640
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
            TabIndex        =   28
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
            PicNormal       =   "ItemsIndex.frx":0950
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
            TabIndex        =   30
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
            PicNormal       =   "ItemsIndex.frx":0EEA
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
            TabIndex        =   33
            Top             =   1950
            Width           =   1140
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
            TabIndex        =   32
            Top             =   900
            Width           =   1140
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
            TabIndex        =   31
            Top             =   1425
            Width           =   1140
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
            TabIndex        =   29
            Top             =   900
            Width           =   4365
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
            TabIndex        =   12
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
            Left            =   1575
            Top             =   1050
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
            TabIndex        =   10
            Top             =   2625
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
            TabIndex        =   8
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
            Height          =   540
            Index           =   0
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   8640
         End
      End
      Begin iGrid300_10Tec.iGrid grdItemsIndex 
         Height          =   7290
         Left            =   75
         TabIndex        =   5
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   1125
         Width           =   6690
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ευρετήριο ειδών"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   30
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   720
         Left            =   75
         TabIndex        =   7
         Top             =   75
         Width           =   3915
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
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   2550
         TabIndex        =   6
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
Attribute VB_Name = "ItemsIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngRowCount As Long
Dim blnError As Boolean
Dim blnProcessing As Boolean


Private Function EditRecord()

    If Items.SeekRecord(grdItemsIndex.CellValue(grdItemsIndex.CurRow, "ID"), txtTable.text) Then
        If Items.Visible Then
            Unload Me
        Else
            With Items
                .Tag = "True"
                .txtTable.text = txtTable.text
                .Show 1
            End With
        End If
    End If

End Function

Private Function FindRecordsAndPopulateGrid()

    If RefreshList > 0 Then
        UpdateRecordCount lblRecordCount, lngRowCount
        UpdateCriteriaLabels lblCategoryDescription.Caption, txtManufacturerDescription.text, txtItemDescription.text
        EnableGrid grdItemsIndex, False
        HighlightRow grdItemsIndex, 1, "", True
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
    
    Open strUnicodeFile For Output As #1
    InitReport myPrinterType, myEAFDSSString, myInvoiceHeight
    GoSub Headers
    
    With grdItemsIndex
        For lngRow = 1 To .RowCount
            Print #1, Tab(1); .CellText(lngRow, "CategoryDescription"); Tab(42); .CellText(lngRow, "ManufacturerDescription"); Tab(83); .CellText(lngRow, "ItemDescription"); Tab(133); Left(.CellText(lngRow, "ItemActive"), 44)
            intProcessedDetailLines = intProcessedDetailLines + 1
            If intProcessedDetailLines > Val(myDetailLines) Then
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
    PrintColumnHeadings 1, "ΚΑΤΗΓΟΡΙΑ", 42, "ΚΑΤΑΣΚΕΥΑΣΤΗΣ", 83, "ΠΕΡΙΓΡΑΦΗ", 131, "ΕΝΕΡΓΟ"
    Print #1, ""
    intProcessedDetailLines = 6
    
    Return
    
ErrTrap:
    Close #1
    CreateUnicodeFile = "Error"
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function UpdateCriteriaLabels(Category, Manufacturer, Description)

    Dim strCriteriaA As String
    Dim strCriteriaB As String

    If Category = "" Then
        strCriteriaA = "Κατηγορία [ ΟΛΕΣ ] "
    Else
        strCriteriaA = "Κατηγορία [ " & Category & " ] "
    End If
    
    If Manufacturer = "" Then
        strCriteriaA = strCriteriaA & "Κατασκευαστής [ ΟΛΟΙ ] "
    Else
        strCriteriaA = strCriteriaA & "Κατασκευαστής [ " & Manufacturer & " ] "
    End If
    
    If Description = "" Then
        strCriteriaB = "Είδη [ ΟΛΑ ]"
    Else
        If Left(Description, 1) <> "*" Then
            strCriteriaB = "Περιγραφή αρχίζει από [ " & UCase(Description) & " ]"
        End If
        If Left(Description, 1) = "*" Then
            strCriteriaB = "Περιγραφή περιέχει το [ " & UCase(Right(Description, Len(Description) - 1)) & " ]"
        End If
    End If
    
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
            AbortProcedure False
        Case 5
            AbortProcedure True
    End Select
    
End Sub

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function
    
    If Not blnStatus Then
        ClearFields grdItemsIndex, lblRecordCount, lblCriteria, lblSelected
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

    'Αρχικές τιμές
    intIndex = 0
    lngRowCount = 0
    frmCriteria(0).Visible = False
    Set TempQuery = CommonDB.CreateQueryDef("")

    'Πλέγμα
    With grdItemsIndex
        .Clear
        .Editable = False
        .Redraw = False
        .RowMode = False
    End With
    
    'Κυρίως διαδικασία
    strSQL = "SELECT Categories.CategoryDescription, Items.ItemID, Items.ItemDescription, ItemActive, Manufacturers.ManufacturerDescription " _
    & "FROM (Items INNER JOIN Categories ON Items.ItemCategoryID = Categories.CategoryID) " _
    & "INNER JOIN Manufacturers ON Items.ItemManufacturerID = Manufacturers.ManufacturerID "

    'Κατηγορία
    If txtCategoryID.text <> "" Then
        strThisParameter = "lngCategoryID Long"
        strThisQuery = "ItemCategoryID = lngCategoryID"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtCategoryID.text)
    End If
    
    'Κατασκευαστής
    If txtManufacturerID.text <> "" Then
        strThisParameter = "lngManufacturerID Long"
        strThisQuery = "ItemManufacturerID = lngManufacturerID"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtManufacturerID.text)
    End If
    
    'Περιγραφή
    If txtItemDescription.text <> "" Then
        strThisParameter = "strDescription String"
        If Left(txtItemDescription.text, 1) <> "*" Then
            strThisQuery = "Left(ItemQuickDescription, Len(strDescription))= strDescription"
        End If
        If Left(txtItemDescription.text, 1) = "*" Then
            strThisQuery = "InStr(ItemQuickDescription, strDescription)"
        End If
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = IIf(Left(txtItemDescription.text, 1) <> "*", txtItemDescription.text, Right(txtItemDescription.text, Len(txtItemDescription.text) - 1))
    End If
    
    'Ταξινόμηση
    strOrder = " ORDER BY CategoryDescription, ManufacturerDescription, ItemDescription"
    
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
    UpdateButtons Me, 5, 0, 0, 0, 0, 1, 0
    cmdButton(4).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        grdItemsIndex.AddRow , , , , , , , rstRecordset.RecordCount
        Do While Not .EOF
            lngRowCount = lngRowCount + 1
            grdItemsIndex.CellValue(lngRowCount, "ID") = !ItemID
            grdItemsIndex.CellValue(lngRowCount, "CategoryDescription") = !CategoryDescription
            grdItemsIndex.CellValue(lngRowCount, "ManufacturerDescription") = !ManufacturerDescription
            grdItemsIndex.CellValue(lngRowCount, "ItemDescription") = !ItemDescription
            grdItemsIndex.CellValue(lngRowCount, "ItemActive") = IIf(!ItemActive = "1", "N", "")
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
        ClearFields grdItemsIndex
    Else
        RefreshList = rstRecordset.RecordCount
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
    
ErrTrap:
    blnError = True
    ClearFields grdItemsIndex, frmProgress
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
            'Κατασκευαστής
            If txtManufacturerDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtManufacturerDescription.text, "Manufacturers", "ManufacturerDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtManufacturerID.text = tmpTableData.strCode
            txtManufacturerDescription.text = tmpTableData.strOneField
    End Select

End Sub

Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdItemsIndex, 44, GetSetting(strAppTitle, "Layout Strings", "grdItemsIndex"), _
            "05NCNID,40NLNCategoryDescription,40NLNManufacturerDescription,50NLNItemDescription,03NCNItemActive,03NCNSelected", _
            "ID,Κατηγορία,Κατασκευαστής,Περιγραφή,Ενεργό,E"
        Me.Refresh
        frmCriteria(0).Visible = True
        txtCategoryShortDescription.SetFocus
    End If
    
    'AddDummyLines grdItemsIndex, 5, 50, 50, 60, 5, 3

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

    SetUpGrid lstIconList, grdItemsIndex
    PositionControls Me, True, grdItemsIndex
    ColorizeControls Me, True
    ClearFields txtCategoryID, txtCategoryShortDescription, lblCategoryDescription, txtManufacturerID, txtManufacturerDescription, txtItemDescription, lblRecordCount, lblCriteria, lblSelected
    UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1

End Sub

Private Sub grdItemsIndex_ColHeaderMouseEnter(ByVal lCol As Long)

    grdItemsIndex.Header.Buttons = True

End Sub

Private Sub grdItemsIndex_ColHeaderMouseLeave(ByVal lCol As Long)

    grdItemsIndex.Header.Buttons = False
    
End Sub

Private Sub grdItemsIndex_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1
    
End Sub

Private Sub grdItemsIndex_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdItemsIndex_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeyInsert Or KeyCode = vbKeyDelete Or KeyCode = vbKeySpace Then
        grdItemsIndex.CellIcon(grdItemsIndex.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdItemsIndex, KeyCode, grdItemsIndex.CurRow, "ID"))
        lblSelected.Caption = CountSelected(grdItemsIndex)
    End If

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdItemsIndex", grdItemsIndex.LayoutCol

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


Private Sub txtManufacturerDescription_Change()

    If txtManufacturerDescription.text = "" Then ClearFields txtManufacturerID

End Sub


Private Sub txtManufacturerDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1

End Sub


Private Sub txtManufacturerDescription_Validate(Cancel As Boolean)

    If txtManufacturerID.text = "" And txtManufacturerDescription.text <> "" Then cmdIndex_Click 1: If txtManufacturerID.text = "" Then Cancel = True

End Sub


