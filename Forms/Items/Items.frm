VERSION 5.00
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form Items 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   9975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16650
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   16650
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   450
      TabIndex        =   27
      Top             =   4725
      Width           =   10365
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   28
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
      End
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   1
         Left            =   1650
         TabIndex        =   29
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
         TabIndex        =   30
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
         TabIndex        =   31
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
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         Caption         =   "Εύρεση"
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
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         Caption         =   "Καρτέλα"
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
         TabIndex        =   34
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
      Height          =   2640
      Left            =   8850
      TabIndex        =   12
      Top             =   1500
      Width           =   4515
      Begin VB.TextBox txtQuickDescription 
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
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "4"
         Top             =   1350
         Width           =   780
      End
      Begin VB.TextBox Text4 
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
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "Items.ItemQuickDescription"
         Top             =   1350
         Width           =   3540
      End
      Begin VB.TextBox Text5 
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
         Text            =   "Items.ItemActive"
         Top             =   1725
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
         Text            =   "3"
         Top             =   975
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
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "7"
         Top             =   2100
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
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Table"
         Top             =   2100
         Width           =   3540
      End
      Begin VB.TextBox txtActiveID 
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
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "5"
         Top             =   1725
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
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "Items.CategoryID"
         Top             =   600
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
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "2"
         Top             =   600
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
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Items.ItemManufacturerID"
         Top             =   975
         Width           =   3540
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
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "Items.ItemID"
         Top             =   225
         Width           =   3540
      End
      Begin VB.TextBox txtID 
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
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   225
         Width           =   780
      End
   End
   Begin UserControls.newText txtDescription 
      Height          =   465
      Left            =   2025
      TabIndex        =   2
      Top             =   2175
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
   Begin UserControls.newText txtCategoryShortDescription 
      Height          =   465
      Left            =   2025
      TabIndex        =   0
      Top             =   1125
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
   Begin UserControls.newText txtManufacturerDescription 
      Height          =   465
      Left            =   2025
      TabIndex        =   1
      Top             =   1650
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1650
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
      PicNormal       =   "Items.frx":0000
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   4
      Left            =   7500
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1650
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
      PicNormal       =   "Items.frx":059A
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   0
      Left            =   2700
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1125
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
      PicNormal       =   "Items.frx":0B34
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   3
      Left            =   3150
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1125
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
      PicNormal       =   "Items.frx":10CE
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newFloat mskVATPercent 
      Height          =   465
      Left            =   2025
      TabIndex        =   4
      Top             =   3225
      Width           =   1440
      _ExtentX        =   2540
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
   Begin UserControls.newFloat mskBalance 
      Height          =   465
      Left            =   2025
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3750
      Width           =   1440
      _ExtentX        =   2540
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
   Begin UserControls.newText txtActiveDescription 
      Height          =   465
      Left            =   2025
      TabIndex        =   3
      Top             =   2700
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   0
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
      Index           =   2
      Left            =   2700
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2700
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
      PicNormal       =   "Items.frx":1668
      PicSizeH        =   16
      PicSizeW        =   16
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
      Left            =   3600
      TabIndex        =   39
      Top             =   1200
      Width           =   4365
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   10800
      Top             =   4425
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   3600
      Top             =   5400
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   2400
      Top             =   4200
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   0
      Left            =   1575
      Top             =   1425
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
      Left            =   2100
      Top             =   0
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
      Top             =   1575
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Είδη"
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
      TabIndex        =   11
      Top             =   75
      Width           =   1050
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Ποσοστό Φ.Π.Α."
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
      Index           =   8
      Left            =   450
      TabIndex        =   10
      Top             =   3300
      Width           =   1140
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000040C0&
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
      TabIndex        =   9
      Top             =   1200
      Width           =   1140
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Ενεργό"
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
      TabIndex        =   8
      Top             =   2775
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
      TabIndex        =   7
      Top             =   1725
      Width           =   1140
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   450
      TabIndex        =   6
      Top             =   3825
      Width           =   1140
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
      TabIndex        =   5
      Top             =   2250
      Width           =   1140
   End
   Begin VB.Shape shpBackground 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   -75
      Top             =   0
      Width           =   840
   End
   Begin VB.Menu mnuHdrPopUp 
      Caption         =   "mnuHdrPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuΑποθήκευσηΠλάτουςΣτηλών 
         Caption         =   "Αποθήκευση πλάτους στηλών"
      End
   End
End
Attribute VB_Name = "Items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim blnStatus As Boolean

Private Function AbortProcedure(blnStatus)
    
    If Not blnStatus Then
        If MyMsgBox(3, strAppTitle, strMessages(3), 2) Then
            blnStatus = False
            ClearFields txtID, txtDescription, txtCategoryID, txtCategoryShortDescription, lblCategoryDescription, txtManufacturerID, txtManufacturerDescription, txtQuickDescription, txtActiveID, txtActiveDescription, mskVATPercent, mskBalance
            DisableFields txtCategoryShortDescription, txtManufacturerDescription, txtDescription, txtActiveDescription, mskVATPercent, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4)
            UpdateButtons Me, 6, 1, 0, 0, IIf(CheckForLoadedForm("ItemsIndex"), 0, 1), 0, 0, 1
        End If
        Exit Function
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Function CheckDuplicateTaxNo(tmpTaxNo)

    Dim tmpRecordset As Recordset
    Dim tmpTableData As typTableData
    
    Set tmpRecordset = CheckForMatch("CommonDB", tmpTaxNo, txtTable.text, "TaxNo", "String", 1, True)
    
    CheckDuplicateTaxNo = ""
    
    If blnStatus Then
        If tmpRecordset.RecordCount = 1 Then
            CheckDuplicateTaxNo = tmpRecordset.Fields(1)
        End If
    End If
    
    If Not blnStatus Then
        If tmpRecordset.RecordCount = 1 And tmpRecordset.Fields(0) <> txtID.text Then
            CheckDuplicateTaxNo = tmpRecordset.Fields(1)
        End If
    End If

End Function

Private Function CreateQuickDescription(myDescription)

    Dim intLoop As Integer
    
    For intLoop = 1 To Len(myDescription)
        If Asc(Mid(myDescription, intLoop, 1)) >= 48 And Asc(Mid(myDescription, intLoop, 1)) <= 57 Then
            CreateQuickDescription = CreateQuickDescription & Mid(myDescription, intLoop, 1)
        End If
    Next intLoop
    
    If Len(CreateQuickDescription) = 0 Then
        CreateQuickDescription = myDescription
    End If
    
End Function

Public Function SeekRecord(myID, myTable)

    On Error GoTo ErrTrap
    
    Dim blnEnableDelete As Boolean
    Dim tmpRecordset As Recordset
    Dim tmpTableData As typTableData
    
    ClearFields txtID, txtDescription, txtCategoryID, txtCategoryShortDescription, lblCategoryDescription, txtManufacturerID, txtManufacturerDescription, txtQuickDescription, txtActiveID, txtActiveDescription, mskVATPercent, mskBalance
    DisableFields txtCategoryShortDescription, txtManufacturerDescription, txtDescription, txtActiveDescription, mskVATPercent, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4)
    
    SeekRecord = False
    
    blnEnableDelete = SimpleSeek("InvoicesTrn", "ItemID", myID)
    
    If MainSeekRecord("CommonDB", myTable, "ID", myID, True, txtID, txtCategoryID, txtManufacturerID, txtDescription, txtQuickDescription, txtActiveID, mskVATPercent) Then
        'Κατηγορία
        Set tmpRecordset = CheckForMatch("CommonDB", txtCategoryID.text, "Categories", "CategoryID", "Numeric", 0, 1)
        txtCategoryID.text = tmpRecordset.Fields(0)
        txtCategoryShortDescription.text = tmpRecordset.Fields(1)
        lblCategoryDescription.Caption = tmpRecordset.Fields(2)
        'Κατασκευαστής
        Set tmpRecordset = CheckForMatch("CommonDB", txtManufacturerID.text, "Manufacturers", "ManufacturerID", "Numeric", 0, 1)
        txtManufacturerID.text = tmpRecordset.Fields(0)
        txtManufacturerDescription.text = tmpRecordset.Fields(1)
        'Ενεργό
        Set tmpRecordset = CheckForMatch("CommonDB", txtActiveID.text, "YesOrNo", "YesNoID", "Numeric", 0, 1)
        txtActiveID.text = tmpRecordset.Fields(0)
        txtActiveDescription.text = tmpRecordset.Fields(1)
        '
        EnableFields txtCategoryShortDescription, txtManufacturerDescription, txtDescription, txtActiveDescription, mskVATPercent, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4)
        UpdateButtons Me, 6, 0, 1, IIf(blnEnableDelete, 1, 0), 0, 1, 1, 0
        blnStatus = False
        SeekRecord = True
    End If
    
    Exit Function
    
ErrTrap:
    SeekRecord = False
    DisplayErrorMessage True, Err.Description

End Function

Private Function DeleteRecord()
    
    If MainDeleteRecord("CommonDB", txtTable.text, strAppTitle, "ID", txtID.text, "True") Then
        ClearFields txtID, txtDescription, txtCategoryID, txtCategoryShortDescription, lblCategoryDescription, txtManufacturerID, txtManufacturerDescription, txtQuickDescription, txtActiveID, txtActiveDescription, mskVATPercent, mskBalance
        DisableFields txtCategoryShortDescription, txtManufacturerDescription, txtDescription, txtActiveDescription, mskVATPercent, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4)
        UpdateButtons Me, 6, 1, 0, 0, IIf(CheckForLoadedForm("ItemsIndex"), 0, 1), 0, 0, 1
    End If

End Function

Private Function NewRecord()
    
    Dim tmpRecordset As Recordset
    
    blnStatus = True
    ClearFields txtID, txtDescription, txtCategoryID, txtCategoryShortDescription, lblCategoryDescription, txtManufacturerID, txtManufacturerDescription, txtQuickDescription, txtActiveID, txtActiveDescription, mskVATPercent, mskBalance
    EnableFields txtCategoryShortDescription, txtManufacturerDescription, txtDescription, txtActiveDescription, mskVATPercent, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4)
    InitializeFields mskVATPercent
    UpdateButtons Me, 6, 0, 1, 0, 0, 0, 1, 0
    txtCategoryShortDescription.SetFocus
    'Ενεργό
    txtActiveID.text = "1"
    Set tmpRecordset = CheckForMatch("CommonDB", txtActiveID.text, "YesOrNo", "YesNoID", "Numeric", 0, 1)
    txtActiveID.text = tmpRecordset.Fields(0)
    txtActiveDescription.text = tmpRecordset.Fields(1)
    'Ποσοστό ΦΠΑ
    mskVATPercent.text = format(curExtraChargesVATPercent, "#,#0.00")

End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    txtQuickDescription.text = CreateQuickDescription(txtDescription.text)
    
    lngItemID = MainSaveRecord("CommonDB", txtTable.text, blnStatus, strAppTitle, "ID", txtID.text, txtCategoryID.text, txtManufacturerID.text, txtDescription.text, txtQuickDescription.text, txtActiveID.text, mskVATPercent.text, 0, 0, 0, 1, strCurrentUser)
             
    If lngItemID <> 0 Then
        ClearFields txtID, txtDescription, txtCategoryID, txtCategoryShortDescription, lblCategoryDescription, txtManufacturerID, txtManufacturerDescription, txtQuickDescription, txtActiveID, txtActiveDescription, mskVATPercent, mskBalance
        DisableFields txtCategoryShortDescription, txtManufacturerDescription, txtDescription, txtActiveDescription, mskVATPercent, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4)
        UpdateButtons Me, 6, 1, 0, 0, IIf(CheckForLoadedForm("ItemsIndex"), 0, 1), 0, 0, 1
    Else
        DisplayErrorMessage True, Err.Description
    End If
    
End Function

Private Function ValidateFields()

    ValidateFields = False
    
    'Κατηγορία
    If DisplayMessage(1, 4, 1, "", txtCategoryID.text) Then txtCategoryShortDescription.SetFocus: Exit Function
    
    'Κατασκευαστής
    If DisplayMessage(1, 4, 1, "", txtManufacturerID.text) Then txtManufacturerDescription.SetFocus: Exit Function
    
    'Περιγραφή
    If DisplayMessage(1, 4, 1, "", txtDescription.text) Then txtDescription.SetFocus: Exit Function
    
    'Ενεργό
    If DisplayMessage(1, 4, 1, "", txtActiveID.text) Then txtActiveDescription.SetFocus: Exit Function
    
    'Φ.Π.Α.
    If DisplayMessage(1, 4, 1, "", mskVATPercent.text) Then mskVATPercent.SetFocus: Exit Function
    
    ValidateFields = True

End Function

Private Function CheckTaxNo(tmpTaxNo)

    On Error GoTo ErrTrap
    
    Dim intLoop As Integer
    Dim lngSum As Long
    Dim lngRemainder As Long
    
    CheckTaxNo = True
    
    If Len(tmpTaxNo) <> 9 Then
        CheckTaxNo = False
        Exit Function
    End If
    
    For intLoop = 1 To Len(tmpTaxNo)
        If Asc(Mid(tmpTaxNo, intLoop, 1)) < 48 Or Asc(Mid(tmpTaxNo, intLoop, 1)) > 57 Then
            CheckTaxNo = False
            Exit Function
        End If
    Next intLoop
    
    lngSum = 256 * Mid(tmpTaxNo, 1, 1) + 128 * Mid(tmpTaxNo, 2, 1) + 64 * Mid(tmpTaxNo, 3, 1) + 32 * Mid(tmpTaxNo, 4, 1) + 16 * Mid(tmpTaxNo, 5, 1) + 8 * Mid(tmpTaxNo, 6, 1) + 4 * Mid(tmpTaxNo, 7, 1) + 2 * Mid(tmpTaxNo, 8, 1)
    
    lngRemainder = lngSum Mod 11
    
    If lngRemainder = 10 Then
        lngRemainder = 0
    End If
    
    If Val(Right(tmpTaxNo, 1)) <> lngRemainder Then
        CheckTaxNo = False
    End If
    
    Exit Function
    
ErrTrap:
    If Err.Number = 13 Then
        CheckTaxNo = False
        Exit Function
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
            ShowIndex
        Case 4
            ShowLedger
        Case 5
            AbortProcedure False
        Case 6
            AbortProcedure True
    End Select

End Sub

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
        Case 2
            'Ενεργό
            If txtActiveDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtActiveDescription.text, "YesOrNo", "YesNoDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtActiveID.text = tmpTableData.strCode
            txtActiveDescription.text = tmpTableData.strOneField
        Case 3
            'Κατηγορία
            With UtilsItemCategories
                .Tag = "True"
                .Show 1, Me
            End With
        Case 4
            'Κατασκευαστής
            With UtilsManufacturers
                .Tag = "True"
                .Show 1, Me
            End With
    End Select

End Sub

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
        Case vbKeyInsert And cmdButton(0).Enabled, vbKeyN And CtrlDown = 4 And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyF10 And cmdButton(1).Enabled, vbKeyS And CtrlDown = 4 And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyF3 And cmdButton(2).Enabled, vbKeyD And CtrlDown = 4 And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyF7 And cmdButton(3).Enabled, vbKeyF And CtrlDown = 4 And cmdButton(3).Enabled
            cmdButton_Click 3
        Case vbKeyF4 And cmdButton(4).Enabled, vbKeyL And CtrlDown = 4 And cmdButton(4).Enabled
            cmdButton_Click 4
        Case vbKeyEscape
            If cmdButton(5).Enabled Then cmdButton_Click 5: Exit Function
            If cmdButton(6).Enabled Then cmdButton_Click 6
        Case vbKeyF12
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()
    
    PositionControls Me, False: ColorizeControls Me
    ClearFields txtID, txtDescription, txtCategoryID, txtCategoryShortDescription, lblCategoryDescription, txtManufacturerID, txtManufacturerDescription, txtQuickDescription, txtActiveID, txtActiveDescription, mskVATPercent, mskBalance
    DisableFields txtCategoryShortDescription, txtManufacturerDescription, txtDescription, txtActiveDescription, mskVATPercent, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4)
    UpdateButtons Me, 6, 1, 0, 0, IIf(CheckForLoadedForm("ItemsIndex"), 0, 1), 0, 0, 1

End Sub

Private Function ShowIndex()

    With ItemsIndex
        .lblTitle.Caption = WindowTitle(lblTitle.Caption)
        .Tag = "True"
        .txtTable.text = txtTable.text
        .Show 1, Me
    End With

End Function

Private Function ShowLedger()

    With ItemsLedger
        .txtCategoryID.text = txtCategoryID.text
        .txtCategoryShortDescription.text = txtCategoryShortDescription.text
        .lblCategoryDescription.Caption = lblCategoryDescription.Caption
        .txtManufacturerID.text = txtManufacturerID.text
        .txtManufacturerDescription.text = txtManufacturerDescription.text
        .txtItemID.text = txtID.text
        .txtItemDescription.text = txtDescription.text
        .txtTable.text = txtTable.text
        .Tag = "True"
        DisableFields .txtCategoryShortDescription, .txtManufacturerDescription, .txtItemDescription, .cmdIndex(0), .cmdIndex(1), .cmdIndex(2)
        .Show 1, Me
    End With

End Function

Private Sub txtActiveDescription_Change()

    If txtActiveDescription.text = "" Then ClearFields txtActiveID

End Sub

Private Sub txtActiveDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2

End Sub


Private Sub txtActiveDescription_Validate(Cancel As Boolean)

    If txtActiveID.text = "" And txtActiveDescription.text <> "" Then cmdIndex_Click 2: If txtActiveID.text = "" Then Cancel = True

End Sub

Private Sub txtCategoryShortDescription_Change()

    If txtCategoryShortDescription.text = "" Then ClearFields txtCategoryID, lblCategoryDescription

End Sub


Private Sub txtCategoryShortDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0
    If KeyCode = vbKeyF5 Then cmdIndex_Click 3

End Sub

Private Sub txtCategoryShortDescription_Validate(Cancel As Boolean)

    If txtCategoryID.text = "" And txtCategoryShortDescription.text <> "" Then cmdIndex_Click 0: If txtCategoryID.text = "" Then Cancel = True
    
End Sub

Private Sub txtManufacturerDescription_Change()

    If txtManufacturerDescription.text = "" Then ClearFields txtManufacturerID

End Sub


Private Sub txtManufacturerDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1
    If KeyCode = vbKeyF5 Then cmdIndex_Click 4

End Sub


Private Sub txtManufacturerDescription_Validate(Cancel As Boolean)

    If txtManufacturerID.text = "" And txtManufacturerDescription.text <> "" Then cmdIndex_Click 1: If txtManufacturerID.text = "" Then Cancel = True

End Sub


