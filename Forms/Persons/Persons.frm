VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form Persons 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   12225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16650
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12225
   ScaleWidth      =   16650
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   450
      TabIndex        =   46
      Top             =   9450
      Width           =   10365
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   47
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
         TabIndex        =   48
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
         TabIndex        =   49
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
         TabIndex        =   50
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
         TabIndex        =   51
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
         TabIndex        =   52
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
         TabIndex        =   53
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
      Height          =   3690
      Left            =   10125
      TabIndex        =   26
      Top             =   3450
      Width           =   4515
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
         TabIndex        =   60
         TabStop         =   0   'False
         Text            =   "7"
         Top             =   2325
         Width           =   2340
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
         TabIndex        =   59
         TabStop         =   0   'False
         Text            =   "OppositeTable"
         Top             =   2325
         Width           =   1965
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
         Left            =   2100
         TabIndex        =   57
         TabStop         =   0   'False
         Text            =   "5"
         Top             =   1575
         Width           =   2340
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
         Index           =   1
         Left            =   75
         TabIndex        =   56
         TabStop         =   0   'False
         Text            =   "Persons.Active"
         Top             =   1575
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
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "6"
         Top             =   2700
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
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "RefersTo"
         Top             =   2700
         Width           =   1965
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
         Index           =   0
         Left            =   75
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "Persons.CountryID"
         Top             =   1200
         Width           =   1965
      End
      Begin VB.TextBox txtCountryID 
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
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "4"
         Top             =   1200
         Width           =   2340
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
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "5"
         Top             =   1950
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
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "Table"
         Top             =   1950
         Width           =   1965
      End
      Begin VB.TextBox txtVATStateID 
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
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "3"
         Top             =   825
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
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "Persons.VATStateID"
         Top             =   825
         Width           =   1965
      End
      Begin VB.TextBox txtTaxOfficeID 
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
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "2"
         Top             =   450
         Width           =   2340
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
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "Persons.TaxOfficeID"
         Top             =   450
         Width           =   1965
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
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Persons.ID"
         Top             =   75
         Width           =   1965
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
         Left            =   2100
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   75
         Width           =   2340
      End
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   75
         Top             =   3075
         _ExtentX        =   953
         _ExtentY        =   953
         Size            =   2296
         Images          =   "Persons.frx":0000
         Version         =   131072
         KeyCount        =   2
         Keys            =   "ÿ"
      End
   End
   Begin UserControls.newText txtDescription 
      Height          =   465
      Left            =   2625
      TabIndex        =   2
      Top             =   1125
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
   Begin UserControls.newText txtPhones 
      Height          =   465
      Left            =   2625
      TabIndex        =   8
      Top             =   4275
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
   Begin UserControls.newText txtAddress 
      Height          =   465
      Left            =   2625
      TabIndex        =   6
      Top             =   3225
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
      Left            =   2625
      TabIndex        =   3
      Top             =   1650
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   0
      MaxLength       =   12
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
   Begin UserControls.newText txtVATStateDescription 
      Height          =   465
      Left            =   2625
      TabIndex        =   10
      Top             =   5325
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
   Begin UserControls.newText txtTaxOfficeDescription 
      Height          =   465
      Left            =   2625
      TabIndex        =   4
      Top             =   2175
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
   Begin UserControls.newText txtCity 
      Height          =   465
      Left            =   2625
      TabIndex        =   7
      Top             =   3750
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
   Begin UserControls.newText txtPersonInCharge 
      Height          =   465
      Left            =   2625
      TabIndex        =   9
      Top             =   4800
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
   Begin UserControls.newText txtProfession 
      Height          =   465
      Left            =   2625
      TabIndex        =   5
      Top             =   2700
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
   Begin UserControls.newText txtEmail 
      Height          =   465
      Left            =   2625
      TabIndex        =   12
      Top             =   6375
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
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   0
      Left            =   7650
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2175
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
      PicNormal       =   "Persons.frx":0918
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   1
      Left            =   8100
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2175
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
      PicNormal       =   "Persons.frx":0EB2
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   2
      Left            =   7650
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5325
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
      PicNormal       =   "Persons.frx":144C
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newFloat mskBalance 
      Height          =   465
      Left            =   2625
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   8475
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
   Begin UserControls.newText txtCountryDescription 
      Height          =   465
      Left            =   2625
      TabIndex        =   14
      Top             =   7425
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
      Index           =   3
      Left            =   7650
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   7425
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
      PicNormal       =   "Persons.frx":19E6
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   4
      Left            =   8100
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   7425
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
      PicNormal       =   "Persons.frx":1F80
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newText txtActiveDescription 
      Height          =   465
      Left            =   2625
      TabIndex        =   15
      Top             =   7950
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
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
      Index           =   5
      Left            =   3300
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   7950
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
      PicNormal       =   "Persons.frx":251A
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newText txtBankAccounts 
      Height          =   465
      Left            =   2625
      TabIndex        =   13
      Top             =   6900
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
   Begin UserControls.newText txtVATArticleNo 
      Height          =   465
      Left            =   2625
      TabIndex        =   11
      Top             =   5850
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   0
      MaxLength       =   5
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
      Caption         =   "Αρθρο απαλλαγής Φ.Π.Α."
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
      Index           =   14
      Left            =   450
      TabIndex        =   61
      Top             =   5925
      Width           =   1815
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Τραπεζικοί λογαριασμοί"
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
      Index           =   13
      Left            =   450
      TabIndex        =   58
      Top             =   6975
      Width           =   1740
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Ενεργός"
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
      Index           =   12
      Left            =   450
      TabIndex        =   55
      Top             =   8025
      Width           =   615
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Χώρα"
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
      Index           =   11
      Left            =   450
      TabIndex        =   40
      Top             =   7500
      Width           =   1740
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
      Index           =   10
      Left            =   450
      TabIndex        =   38
      Top             =   8550
      Width           =   1740
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   10800
      Top             =   8625
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   3600
      Top             =   10125
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   3000
      Top             =   8925
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
      Left            =   2175
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
      Left            =   5550
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
      Top             =   2100
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "E-mail"
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
      Index           =   1
      Left            =   450
      TabIndex        =   25
      Top             =   6450
      Width           =   465
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Συναλλασόμενοι"
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
      TabIndex        =   24
      Top             =   75
      Width           =   3810
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Καθεστώς Φ.Π.Α."
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
      Index           =   9
      Left            =   450
      TabIndex        =   23
      Top             =   5400
      Width           =   1740
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Πόλη"
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
      TabIndex        =   22
      Top             =   3825
      Width           =   1740
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Οικονομική υπηρεσία"
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
      TabIndex        =   21
      Top             =   2250
      Width           =   1740
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
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
      TabIndex        =   20
      Top             =   1725
      Width           =   1740
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Διεύθυνση"
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
      TabIndex        =   19
      Top             =   3300
      Width           =   1740
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Δραστηριότητα"
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
      TabIndex        =   18
      Top             =   2775
      Width           =   1740
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Τηλέφωνα"
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
      TabIndex        =   17
      Top             =   4350
      Width           =   1740
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Υπεύθυνος επικοινωνίας"
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
      Top             =   4875
      Width           =   1740
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
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
      TabIndex        =   1
      Top             =   1200
      Width           =   1740
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
Attribute VB_Name = "Persons"
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
            ClearFields txtID, txtDescription, txtProfession, txtAddress, txtCity, txtPhones, txtPersonInCharge, txtEmail, txtBankAccounts, txtTaxNo, txtTaxOfficeID, txtTaxOfficeDescription, txtVATStateID, txtVATStateDescription, , txtCountryID, txtCountryDescription, txtActiveID, txtActiveDescription, txtVATArticleNo
            DisableFields txtDescription, txtProfession, txtAddress, txtCity, txtPhones, txtPersonInCharge, txtEmail, txtBankAccounts, txtTaxNo, txtTaxOfficeDescription, txtVATStateDescription, txtCountryDescription, txtActiveDescription, txtVATArticleNo, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
            UpdateButtons Me, 6, 1, 0, 0, IIf(CheckForLoadedForm("PersonsIndex"), 0, 1), 0, 0, 1
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
        If tmpRecordset.RecordCount = 1 Then
            If tmpRecordset.Fields(0) <> txtID.text Then
                CheckDuplicateTaxNo = tmpRecordset.Fields(1)
            End If
        End If
    End If

End Function

Public Function SeekRecord(myID, myTable, myRefersTo)

    On Error GoTo ErrTrap
    
    Dim blnEnableDelete As Boolean
    Dim tmpRecordset As Recordset
    Dim tmpTableData As typTableData
    
    ClearFields txtID, txtDescription, txtProfession, txtAddress, txtCity, txtPhones, txtPersonInCharge, txtEmail, txtBankAccounts, txtTaxNo, txtTaxOfficeID, txtTaxOfficeDescription, txtVATStateID, txtVATStateDescription, , txtCountryID, txtCountryDescription, txtActiveID, txtActiveDescription, txtVATArticleNo
    DisableFields txtDescription, txtProfession, txtAddress, txtCity, txtPhones, txtPersonInCharge, txtEmail, txtBankAccounts, txtTaxNo, txtTaxOfficeDescription, txtVATStateDescription, txtCountryDescription, txtActiveDescription, txtVATArticleNo, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
    
    SeekRecord = False
    
    blnEnableDelete = SimpleSeek("Invoices", "PersonID", myID, myRefersTo)
    
    If MainSeekRecord("CommonDB", myTable, "ID", myID, True, txtID, txtDescription, txtTaxNo, txtTaxOfficeID, txtProfession, txtAddress, txtCity, txtPhones, txtPersonInCharge, txtVATStateID, txtVATArticleNo, txtEmail, txtBankAccounts, txtCountryID, txtActiveID) Then
        'Οικονομική υπηρεσία
        Set tmpRecordset = CheckForMatch("CommonDB", txtTaxOfficeID.text, "TaxOffices", "TaxOfficeID", "Numeric", 0, 1)
        txtTaxOfficeID.text = tmpRecordset.Fields(0)
        txtTaxOfficeDescription.text = tmpRecordset.Fields(1)
        'Καθεστώς Φ.Π.Α.
        Set tmpRecordset = CheckForMatch("CommonDB", txtVATStateID.text, "VATStates", "VATStateID", "Numeric", 0, 1)
        txtVATStateID.text = tmpRecordset.Fields(0)
        txtVATStateDescription.text = tmpRecordset.Fields(1)
        'Χώρα
        Set tmpRecordset = CheckForMatch("CommonDB", txtCountryID.text, "Countries", "CountryID", "Numeric", 0, 1)
        txtCountryID.text = tmpRecordset.Fields(0)
        txtCountryDescription.text = tmpRecordset.Fields(2)
        'Ενεργός
        Set tmpRecordset = CheckForMatch("CommonDB", txtActiveID.text, "YesOrNo", "YesNoID", "Numeric", 0, 1)
        txtActiveID.text = tmpRecordset.Fields(0)
        txtActiveDescription.text = tmpRecordset.Fields(1)
        '
        EnableFields txtDescription, txtProfession, txtAddress, txtCity, txtPhones, txtPersonInCharge, txtEmail, txtBankAccounts, txtTaxNo, txtTaxOfficeDescription, txtVATStateDescription, txtCountryDescription, txtActiveDescription, txtVATArticleNo, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
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
        ClearFields txtID, txtDescription, txtProfession, txtAddress, txtCity, txtPhones, txtPersonInCharge, txtEmail, txtBankAccounts, txtTaxNo, txtTaxOfficeID, txtTaxOfficeDescription, txtVATStateID, txtVATStateDescription, txtCountryID, txtCountryDescription, txtActiveID, txtActiveDescription, txtVATArticleNo
        DisableFields txtDescription, txtProfession, txtAddress, txtCity, txtPhones, txtPersonInCharge, txtEmail, txtBankAccounts, txtTaxNo, txtTaxOfficeDescription, txtVATStateDescription, txtCountryDescription, txtActiveDescription, txtVATArticleNo, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
        UpdateButtons Me, 6, 1, 0, 0, IIf(CheckForLoadedForm("PersonsIndex"), 0, 1), 0, 0, 1
    End If

End Function

Private Function NewRecord()
    
    Dim tmpRecordset As Recordset
    
    blnStatus = True
    ClearFields txtID, txtDescription, txtProfession, txtAddress, txtCity, txtPhones, txtPersonInCharge, txtEmail, txtBankAccounts, txtTaxNo, txtTaxOfficeID, txtTaxOfficeDescription, txtVATStateID, txtVATStateDescription, , txtCountryID, txtCountryDescription, txtActiveID, txtActiveDescription, txtVATArticleNo
    EnableFields txtDescription, txtProfession, txtAddress, txtCity, txtPhones, txtPersonInCharge, txtEmail, txtBankAccounts, txtTaxNo, txtTaxOfficeDescription, txtVATStateDescription, txtCountryDescription, txtActiveDescription, txtVATArticleNo, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
    UpdateButtons Me, 6, 0, 1, 0, 0, 0, 1, 0
    txtDescription.SetFocus
    'Ενεργός
    txtActiveID.text = "1"
    Set tmpRecordset = CheckForMatch("CommonDB", txtActiveID.text, "YesOrNo", "YesNoID", "Numeric", 0, 1)
    txtActiveID.text = tmpRecordset.Fields(0)
    txtActiveDescription.text = tmpRecordset.Fields(1)

End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    If MainSaveRecord("CommonDB", txtTable.text, blnStatus, strAppTitle, "ID", txtID.text, txtDescription.text, txtTaxNo.text, txtTaxOfficeID.text, txtProfession.text, txtAddress.text, txtCity.text, txtPhones.text, txtPersonInCharge.text, txtVATStateID.text, txtVATArticleNo.text, txtEmail.text, txtBankAccounts.text, txtCountryID.text, txtActiveID.text, 1, strCurrentUser) <> 0 Then
        ClearFields txtID, txtDescription, txtProfession, txtAddress, txtCity, txtPhones, txtPersonInCharge, txtEmail, txtBankAccounts, txtTaxNo, txtTaxOfficeID, txtTaxOfficeDescription, txtVATStateID, txtVATStateDescription, txtCountryID, txtCountryDescription, txtActiveID, txtActiveDescription, txtVATArticleNo
        DisableFields txtDescription, txtProfession, txtAddress, txtCity, txtPhones, txtPersonInCharge, txtEmail, txtBankAccounts, txtTaxNo, txtTaxOfficeDescription, txtVATStateDescription, txtCountryDescription, txtActiveDescription, txtVATArticleNo, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
        UpdateButtons Me, 6, 1, 0, 0, IIf(CheckForLoadedForm("PersonsIndex"), 0, 1), 0, 0, 1
    Else
        DisplayErrorMessage True, Err.Description
    End If
    
End Function

Private Function ValidateFields()

    ValidateFields = False
    
    'Επωνυμία
    If DisplayMessage(1, 4, 1, "", txtDescription.text) Then txtDescription.SetFocus: Exit Function
    
    'Α.Φ.Μ.
    If DisplayMessage(1, 4, 1, "", txtTaxNo.text) Then txtTaxNo.SetFocus: Exit Function
    
    'Ελεγχος Α.Φ.Μ.
    If blnCheckTaxNo Then
        If Not CheckTaxNo(txtTaxNo.text) Then
            If Not MyMsgBox(3, strAppTitle, strMessages(52), 2) Then
                txtTaxNo.SetFocus
                Exit Function
            End If
        End If
    End If
    
    'Α.Φ.Μ. υπάρχει ήδη
    Dim strPerson As String
    strPerson = CheckDuplicateTaxNo(txtTaxNo.text)
    If strPerson <> "" Then
        If Not MyMsgBox(3, strAppTitle, strMessages(53) & strPerson & Chr(13) & "Θέλετε να συνεχίσετε;", 2) Then
            txtTaxNo.SetFocus
            Exit Function
        End If
    End If
    
    'Οικονομική υπηρεσία
    If DisplayMessage(1, 4, 1, "", txtTaxOfficeID.text) Then txtTaxOfficeDescription.SetFocus: Exit Function
    
    'Καθεστώς Φ.Π.Α.
    If DisplayMessage(1, 4, 1, "", txtVATStateID.text) Then txtVATStateDescription.SetFocus: Exit Function
    
    'Χώρα
    If DisplayMessage(1, 4, 1, "", txtCountryID.text) Then txtCountryDescription.SetFocus: Exit Function
    
    'Ενεργό
    If DisplayMessage(1, 4, 1, "", txtActiveID.text) Then txtActiveDescription.SetFocus: Exit Function

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
            ShowPersonLedger txtID.text, txtDescription.text, IIf(txtRefersTo.text = "3", "Καρτέλα προμηθευτή", "Καρτέλα πελάτη"), txtTable.text, txtOppositeTable.text, txtRefersTo.text
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
            'Οικονομική Υπηρεσία
            If txtTaxOfficeDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtTaxOfficeDescription.text, "TaxOffices", "TaxOfficeDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Ονομασία", 0, 40, 1, 0)
            txtTaxOfficeID.text = tmpTableData.strCode
            txtTaxOfficeDescription.text = tmpTableData.strOneField
        Case 1
            'Οικονομική Υπηρεσία
            With UtilsTaxOffices
                .Tag = "True"
                .Show 1, Me
            End With
        Case 2
            'Καθεστώς Φ.Π.Α.
            If txtVATStateDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtVATStateDescription.text, "VATStates", "VATStateDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Ονομασία", 0, 40, 1, 0)
            txtVATStateID.text = tmpTableData.strCode
            txtVATStateDescription.text = tmpTableData.strOneField
        Case 3
            'Χώρα
            If txtCountryDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtCountryDescription.text, "Countries", "CountryDescription", "String", 1, 3)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 3, 0, 1, 2, "ID", "Συντ.", "Ονομασία", 0, 4, 40, 1, 1, 0)
            txtCountryID.text = tmpTableData.strCode
            txtCountryDescription.text = tmpTableData.strTwoField
        Case 4
            'Χώρα
            With UtilsCountries
                .Tag = "True"
                .Show 1, Me
            End With
        Case 5
            'Ενεργός
            If txtActiveDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtActiveDescription.text, "YesOrNo", "YesNoDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtActiveID.text = tmpTableData.strCode
            txtActiveDescription.text = tmpTableData.strOneField
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
        Case vbKeyF12 And CtrlDown = 4
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()
    
    PositionControls Me, False: ColorizeControls Me
    ClearFields txtID, txtDescription, txtProfession, txtAddress, txtCity, txtPhones, txtPersonInCharge, txtEmail, txtBankAccounts, txtTaxNo, txtTaxOfficeID, txtTaxOfficeDescription, txtVATStateID, txtVATStateDescription, , txtCountryID, txtCountryDescription, txtActiveID, txtActiveDescription, txtVATArticleNo
    DisableFields txtDescription, txtProfession, txtAddress, txtCity, txtPhones, txtPersonInCharge, txtEmail, txtBankAccounts, txtTaxNo, txtTaxOfficeDescription, txtVATStateDescription, txtCountryDescription, txtActiveDescription, txtVATArticleNo, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
    UpdateButtons Me, 6, 1, 0, 0, IIf(CheckForLoadedForm("PersonsIndex"), 0, 1), 0, 0, 1
    
End Sub

Private Sub txtActiveDescription_Change()

    If txtActiveDescription.text = "" Then ClearFields txtActiveID

End Sub

Private Sub txtActiveDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 5
    
End Sub


Private Sub txtActiveDescription_Validate(Cancel As Boolean)

    If txtActiveID.text = "" And txtActiveDescription.text <> "" Then cmdIndex_Click 5: If txtActiveID.text = "" Then Cancel = True
    
End Sub

Private Sub txtCountryDescription_Change()

    If txtCountryDescription.text = "" Then ClearFields txtCountryID

End Sub

Private Sub txtCountryDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 3
    If KeyCode = vbKeyF5 Then cmdIndex_Click 4

End Sub


Private Sub txtCountryDescription_Validate(Cancel As Boolean)

    If txtCountryID.text = "" And txtCountryDescription.text <> "" Then cmdIndex_Click 3: If txtCountryID.text = "" Then Cancel = True

End Sub

Private Sub txtTaxOfficeDescription_Change()

    If txtTaxOfficeDescription.text = "" Then ClearFields txtTaxOfficeID
    
End Sub

Private Sub txtTaxOfficeDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0
    If KeyCode = vbKeyF5 Then cmdIndex_Click 1

End Sub

Private Sub txtTaxOfficeDescription_Validate(Cancel As Boolean)

    If txtTaxOfficeID.text = "" And txtTaxOfficeDescription.text <> "" Then cmdIndex_Click 0: If txtTaxOfficeID.text = "" Then Cancel = True

End Sub

Private Sub txtVATStateDescription_Change()

    If txtVATStateDescription.text = "" Then ClearFields txtVATStateID
    
End Sub

Private Sub txtVATStateDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2

End Sub

Private Sub txtVATStateDescription_Validate(Cancel As Boolean)

    If txtVATStateID.text = "" And txtVATStateDescription.text <> "" Then cmdIndex_Click 2: If txtVATStateID.text = "" Then Cancel = True

End Sub

Private Function ShowIndex()

    With PersonsIndex
        .Tag = "True"
        .txtTable.text = txtTable.text
        .txtOppositeTable.text = txtOppositeTable.text
        .txtRefersTo.text = txtRefersTo.text
        .lblTitle.Caption = IIf(txtRefersTo.text = "3", "Ευρετήριο προμηθευτών", "Ευρετήριο πελατών")
        .Show 1, Me
    End With

End Function

