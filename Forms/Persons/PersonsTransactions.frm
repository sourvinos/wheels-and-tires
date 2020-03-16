VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form PersonsTransactions 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   13320
   ClientLeft      =   15
   ClientTop       =   0
   ClientWidth     =   15435
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13320
   ScaleWidth      =   15435
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   825
      TabIndex        =   35
      Top             =   8775
      Width           =   10365
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   36
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
         TabIndex        =   37
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
         TabIndex        =   38
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
         TabIndex        =   39
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
         TabIndex        =   40
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
         TabIndex        =   41
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
         TabIndex        =   42
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
      ForeColor       =   &H00FFFFFF&
      Height          =   3315
      Left            =   9375
      TabIndex        =   6
      Top             =   3375
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
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "OppositeTable"
         Top             =   1950
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
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "6"
         Top             =   1950
         Width           =   2340
      End
      Begin VB.TextBox txtInvoiceInTime 
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
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "4"
         Top             =   1200
         Width           =   2340
      End
      Begin VB.TextBox Text6 
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
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceInTime"
         Top             =   1200
         Width           =   1965
      End
      Begin VB.TextBox txtInvoiceInDate 
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
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceInDate"
         Top             =   825
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
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "5"
         Top             =   1575
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
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Table"
         Top             =   1575
         Width           =   1965
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
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "RefersTo"
         Top             =   2325
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
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "7"
         Top             =   2325
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
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceCodeID"
         Top             =   450
         Width           =   1965
      End
      Begin VB.TextBox txtInvoiceCodeID 
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
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "2"
         Top             =   450
         Width           =   2340
      End
      Begin VB.TextBox txtInvoiceTrnID 
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
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   75
         Width           =   2340
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
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceTrnID"
         Top             =   75
         Width           =   1965
      End
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   75
         Top             =   2700
         _ExtentX        =   953
         _ExtentY        =   953
         IconSizeX       =   26
         IconSizeY       =   32
         Size            =   14064
         Images          =   "PersonsTransactions.frx":0000
         Version         =   131072
         KeyCount        =   4
         Keys            =   "ÿÿÿ"
      End
   End
   Begin VB.Frame frmFrame 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   7140
      Index           =   0
      Left            =   1875
      TabIndex        =   15
      Top             =   1125
      Width           =   10215
      Begin UserControls.newFloat mskTotal 
         Height          =   465
         Left            =   7425
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   6225
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
      Begin UserControls.newDate mskInvoiceIssueDate 
         Height          =   465
         Left            =   2175
         TabIndex        =   1
         Top             =   450
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
      Begin iGrid300_10Tec.iGrid grdPersonsTransactions 
         Height          =   3615
         Left            =   2175
         TabIndex        =   5
         Top             =   2550
         Width           =   6990
         _ExtentX        =   12330
         _ExtentY        =   6376
         Appearance      =   0
         BackColor       =   16777215
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
         FrozenColsEdgeColor=   0
         HighlightBackColor=   12648384
         HighlightForeColor=   0
         HighlightForeColorNoFocus=   0
      End
      Begin UserControls.newText txtInvoiceRemarks 
         Height          =   465
         Left            =   2175
         TabIndex        =   4
         Top             =   2025
         Width           =   6990
         _ExtentX        =   12330
         _ExtentY        =   820
         ForeColor       =   0
         MaxLength       =   60
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
      Begin UserControls.newText txtCodeShortDescription 
         Height          =   465
         Left            =   2175
         TabIndex        =   2
         Top             =   975
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   820
         Alignment       =   2
         ForeColor       =   0
         MaxLength       =   4
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
         Left            =   3750
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   975
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
         PicNormal       =   "PersonsTransactions.frx":3710
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   1
         Left            =   4200
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   975
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
         PicNormal       =   "PersonsTransactions.frx":3CAA
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newText txtInvoiceNo 
         Height          =   465
         Left            =   2175
         TabIndex        =   3
         Top             =   1500
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   820
         Alignment       =   2
         ForeColor       =   0
         MaxLength       =   7
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
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   4
         Left            =   6975
         Top             =   6375
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   465
         Index           =   3
         Left            =   2700
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   465
         Index           =   20
         Left            =   8025
         Top             =   6675
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Ημερομηνία"
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
         TabIndex        =   24
         Top             =   525
         Width           =   1290
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   1140
         Index           =   1
         Left            =   1725
         Top             =   900
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Παρατηρήσεις"
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
         TabIndex        =   23
         Top             =   2100
         Width           =   1290
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Νο παραστατικού"
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
         TabIndex        =   22
         Top             =   1575
         Width           =   1290
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Παραστατικό"
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
         TabIndex        =   21
         Top             =   1050
         Width           =   1290
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Σύνολο"
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
         Left            =   6450
         TabIndex        =   20
         Top             =   6300
         Width           =   540
      End
      Begin VB.Label lblCodeDescription 
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
         Left            =   4650
         TabIndex        =   19
         Top             =   1050
         Width           =   4365
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   5
         Left            =   0
         Top             =   1200
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin Dacara_dcButton.dcButton btnPanel 
      Height          =   990
      Index           =   0
      Left            =   450
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1125
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1746
      BackColor       =   12640511
      ButtonShape     =   3
      ButtonStyle     =   2
      Caption         =   "Στοιχεία"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8388736
      State           =   3
   End
   Begin Dacara_dcButton.dcButton btnPanel 
      Height          =   990
      Index           =   1
      Left            =   450
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2175
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1746
      BackColor       =   12640511
      ButtonShape     =   3
      ButtonStyle     =   2
      Caption         =   "Αξιόγραφα"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8388736
   End
   Begin VB.Frame frmFrame 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   7140
      Index           =   1
      Left            =   4275
      TabIndex        =   27
      Top             =   1125
      Width           =   10215
      Begin iGrid300_10Tec.iGrid grdPersonsTransactionsChecks 
         Height          =   5715
         Left            =   450
         TabIndex        =   28
         Top             =   450
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   10081
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
         FrozenColsEdgeColor=   0
         HighlightBackColor=   12648384
         HighlightForeColor=   0
         HighlightForeColorNoFocus=   0
      End
      Begin UserControls.newFloat mskTotalChecks 
         Height          =   465
         Left            =   8025
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   6225
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   820
         Enabled         =   0   'False
         Alignment       =   1
         ForeColor       =   0
         Text            =   "0,00"
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
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   465
         Index           =   7
         Left            =   8475
         Top             =   6675
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   1140
         Index           =   6
         Left            =   7575
         Top             =   5925
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Σύνολο"
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
         Left            =   7050
         TabIndex        =   30
         Top             =   6300
         Width           =   540
      End
   End
   Begin VB.Shape shpBridge 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1005
      Index           =   0
      Left            =   450
      Top             =   1125
      Width           =   3090
   End
   Begin VB.Shape shpBridge 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1005
      Index           =   1
      Left            =   450
      Top             =   2175
      Width           =   3090
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   3750
      Top             =   9450
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   2400
      Top             =   7350
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   12075
      Top             =   7350
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
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Κινήσεις συναλλασόμενων"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   30
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   720
      Left            =   225
      TabIndex        =   0
      Top             =   75
      Width           =   5805
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
      Left            =   4050
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
Attribute VB_Name = "PersonsTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim blnStatus As Boolean
Dim blnError As Boolean
Dim strGridFocus As String
Dim blnGridEditInProgress As Boolean
Dim lngTrnID As Long

Private Function AbortProcedure(blnStatus)
    
     If blnGridEditInProgress Then
        blnGridEditInProgress = False
        grdPersonsTransactions.CancelEdit
        grdPersonsTransactionsChecks.CancelEdit
        Exit Function
    End If
    
    If Not blnStatus Then
        If MyMsgBox(3, strAppTitle, strMessages(3), 2) Then
            btnPanel_Click 0
            blnStatus = False
            ClearFields txtInvoiceTrnID, txtInvoiceCodeID, txtInvoiceInDate, txtInvoiceInTime, mskInvoiceIssueDate, txtCodeShortDescription, lblCodeDescription, txtInvoiceNo, txtInvoiceRemarks, grdPersonsTransactions, grdPersonsTransactionsChecks, mskTotal, mskTotalChecks
            DisableFields mskInvoiceIssueDate, txtCodeShortDescription, txtInvoiceNo, txtInvoiceRemarks, grdPersonsTransactions, grdPersonsTransactionsChecks, cmdIndex(0), cmdIndex(1), btnPanel(1)
            UpdateButtons Me, 6, 1, 0, 0, IIf(CheckForLoadedForm("CommonTransactionsIndex"), 0, 1), 0, 0, 1
        End If
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Function DeleteChecks()

    On Error GoTo ErrTrap
    
    Dim strSQL As String
    
    If blnError Then Exit Function
        
    strSQL = "DELETE FROM Checks WHERE CheckTrnID = " & Val(txtInvoiceTrnID.text)
    CommonDB.Execute (strSQL)
    
    Exit Function
    
ErrTrap:
    blnError = True
    DeleteChecks = False
    DisplayErrorMessage True, Err.Description

End Function

Private Function DeleteInvoices()

    On Error GoTo ErrTrap
    
    Dim strSQL As String
    
    If blnError Then Exit Function
    
    strSQL = "DELETE FROM Invoices WHERE InvoiceTrnID = " & Val(txtInvoiceTrnID.text)
    CommonDB.Execute (strSQL)
    
    Exit Function
    
ErrTrap:
    blnError = True
    DeleteInvoices = False
    DisplayErrorMessage True, Err.Description

End Function


Private Function DeleteRecord()
    
    On Error GoTo ErrTrap
    
    Dim strSQL As String
    
    If Not MyMsgBox(3, strAppTitle, strMessages(4), 2) Then Exit Function
    
    blnError = False
    
    BeginTrans
    
    DeleteInvoices
    DeleteChecks
    
    If Not blnError Then
        CommitTrans
        btnPanel_Click 0
        ClearFields txtInvoiceTrnID, txtInvoiceCodeID, txtInvoiceInDate, txtInvoiceInTime, mskInvoiceIssueDate, txtCodeShortDescription, lblCodeDescription, txtInvoiceNo, txtInvoiceRemarks, grdPersonsTransactions, grdPersonsTransactionsChecks, mskTotal, mskTotalChecks
        DisableFields mskInvoiceIssueDate, txtCodeShortDescription, txtInvoiceNo, txtInvoiceRemarks, grdPersonsTransactions, grdPersonsTransactionsChecks, cmdIndex(0), cmdIndex(1), btnPanel(1)
        UpdateButtons Me, 6, 1, 0, 0, IIf(CheckForLoadedForm("CommonTransactionsIndex"), 0, 1), 0, 0, 1
    Else
        Rollback
    End If
    
    Exit Function
    
ErrTrap:
    Rollback
    DeleteRecord = False
    DisplayErrorMessage True, Err.Description
    
End Function

Function DoSharedStuff(myInvoiceTrnID, myWindowTitle, myTable, myRefersTo, myOppositeTable)

    FindInvoicesWithTrnID myInvoiceTrnID, myWindowTitle, myTable, myRefersTo, myOppositeTable
    FindChecksWithTrnID myInvoiceTrnID, myOppositeTable
    Me.Tag = "False"
    If Me.Visible Then
        Unload CommonTransactionsIndex
        Me.mskInvoiceIssueDate.SetFocus
    Else
        Me.Show 1
    End If

End Function

Function FindChecksWithTrnID(myInvoiceTrnID, myOppositeTable)

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
    Dim tmpRecordset As Recordset
    
    'Local μεταβλητές
    Dim lngRow As Long
    Dim lngRowsToAdd  As Long
    Dim bytLoop As Byte
    Dim tmpTableData As typTableData
        
    'Αρχικές τιμές
    lngRow = 0
    CustomizeGrid grdPersonsTransactionsChecks
    If grdPersonsTransactionsChecks.RowCount = 0 Then AddGridLines grdPersonsTransactionsChecks, txtRefersTo.text, 16
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Κύριο SQL
    strSQL = "SELECT CheckBankID, CheckNo, CheckExpireDate,  CheckAmount, CheckIssuedByID, BankDescription, Description " _
        & "FROM (Checks " _
        & "INNER JOIN Banks ON Checks.CheckBankID = Banks.BankID) " _
        & "LEFT JOIN " & myOppositeTable & " ON Checks.CheckIssuedByID = " & myOppositeTable & ".ID "
        
    'TrnID επιταγών
    strThisParameter = "lngCheckTrnID Long"
    strThisQuery = "Checks.CheckTrnID = lngCheckTrnID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(myInvoiceTrnID)
        
    'Ταξινόμηση
    strOrder = " ORDER BY CheckID"
        
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
    If rstRecordset.RecordCount = 0 Then blnError = False: FindChecksWithTrnID = True: Exit Function
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        While Not .EOF
            With grdPersonsTransactionsChecks
                lngRow = lngRow + 1
                grdPersonsTransactionsChecks.CellValue(lngRow, "BankID") = rstRecordset!CheckBankID
                grdPersonsTransactionsChecks.CellValue(lngRow, "BankDescription") = rstRecordset!BankDescription
                grdPersonsTransactionsChecks.CellValue(lngRow, "CheckIssuedByID") = rstRecordset!CheckIssuedByID
                grdPersonsTransactionsChecks.CellValue(lngRow, "IssuedByDescription") = rstRecordset!Description
                grdPersonsTransactionsChecks.CellValue(lngRow, "CheckNo") = rstRecordset!CheckNo
                grdPersonsTransactionsChecks.CellValue(lngRow, "CheckExpire") = rstRecordset!CheckExpireDate
                grdPersonsTransactionsChecks.CellValue(lngRow, "CheckAmount") = rstRecordset!CheckAmount
            End With
            .MoveNext
        Wend
    End With
    
    'Τελικές ενέργειες
    CustomizeGrid grdPersonsTransactionsChecks
    EnableGrid grdPersonsTransactionsChecks, True
    mskTotalChecks.text = format(CalculateColumnTotal(grdPersonsTransactionsChecks, "CheckAmount"), "#,##0.00")
    FindChecksWithTrnID = True
    
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
    FindChecksWithTrnID = False
    DisplayErrorMessage True, Err.Description

End Function

Private Function NewRecord()
    
    blnStatus = True
    ClearFields txtInvoiceTrnID, txtInvoiceCodeID, txtInvoiceInDate, txtInvoiceInTime, mskInvoiceIssueDate, txtCodeShortDescription, lblCodeDescription, txtInvoiceNo, txtInvoiceRemarks, grdPersonsTransactions, grdPersonsTransactionsChecks, mskTotal, mskTotalChecks
    EnableFields mskInvoiceIssueDate, txtCodeShortDescription, txtInvoiceNo, txtInvoiceRemarks, grdPersonsTransactions, grdPersonsTransactionsChecks, cmdIndex(0), cmdIndex(1), btnPanel(1)
    EditableFields grdPersonsTransactions, grdPersonsTransactionsChecks
    CustomizeGrid grdPersonsTransactions, grdPersonsTransactionsChecks
    EnableTabStop grdPersonsTransactions, grdPersonsTransactionsChecks
    AddGridLines grdPersonsTransactions, txtRefersTo.text, 9
    AddGridLines grdPersonsTransactionsChecks, txtRefersTo.text, 16
    InitializeFields mskInvoiceIssueDate, mskTotal, mskTotalChecks
    UpdateButtons Me, 6, 0, 1, 0, 0, 0, 1, 0
    mskInvoiceIssueDate.SetFocus

End Function

Private Function SaveChecks()

    Dim lngRow As Long
    
    If blnError Then Exit Function
    
    With grdPersonsTransactionsChecks
        For lngRow = 1 To .RowCount
            If .CellValue(lngRow, "BankID") <> "" Then
                If Not MainSaveRecord("CommonDB", "Checks", True, strAppTitle, "ID", txtInvoiceTrnID.text, .CellValue(lngRow, "BankID"), .CellValue(lngRow, "CheckNo"), .CellText(lngRow, "CheckExpire"), .CellValue(lngRow, "CheckAmount"), IIf(grdPersonsTransactionsChecks.CellValue(lngRow, "CheckIssuedByID") <> "", grdPersonsTransactionsChecks.CellValue(lngRow, "CheckIssuedByID"), 0), Val(txtRefersTo.text), lngTrnID, strCurrentUser) <> 0 Then
                    blnError = True
                End If
            End If
        Next lngRow
    End With
    
    SaveChecks = True
    
End Function

Private Function SaveInvoices()

    Dim lngRow As Long
    
    If blnError Then Exit Function
    
    lngTrnID = IIf(txtInvoiceTrnID.text = "", AddOneToTheLastRecord, txtInvoiceTrnID.text)
    
    For lngRow = 1 To grdPersonsTransactions.RowCount
        If grdPersonsTransactions.CellValue(lngRow, "ID") <> "" Then
            If Not MainSaveRecord("CommonDB", "Invoices", True, strAppTitle, "ID", lngTrnID, mskInvoiceIssueDate.text, Val(txtInvoiceNo.text), txtInvoiceCodeID.text, Val(txtRefersTo.text), 0, 0, 0, 0, 0, 0, 0, grdPersonsTransactions.CellValue(lngRow, "Amount"), lngTrnID, txtInvoiceRemarks.text, "", 6, grdPersonsTransactions.CellValue(lngRow, "ID"), 0, 0, Date, Time, "", "", "", "", 1, 0, strCurrentUser) <> 0 Then
                blnError = True
            End If
        End If
    Next lngRow
    
End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    blnError = False
    
    BeginTrans
    
    DeleteInvoices
    DeleteChecks
    SaveInvoices
    SaveChecks
    
    If Not blnError Then
        CommitTrans
        btnPanel_Click 0
        ClearFields txtInvoiceTrnID, txtInvoiceCodeID, txtInvoiceInDate, txtInvoiceInTime, mskInvoiceIssueDate, txtCodeShortDescription, lblCodeDescription, txtInvoiceNo, txtInvoiceRemarks, grdPersonsTransactions, grdPersonsTransactionsChecks, mskTotal, mskTotalChecks
        DisableFields mskInvoiceIssueDate, txtCodeShortDescription, txtInvoiceNo, txtInvoiceRemarks, grdPersonsTransactions, grdPersonsTransactionsChecks, cmdIndex(0), cmdIndex(1), btnPanel(1)
        UpdateButtons Me, 6, 1, 0, 0, IIf(CheckForLoadedForm("CommonTransactionsIndex"), 0, 1), 0, 0, 1
    Else
        Rollback
    End If
    
End Function

Function FindInvoicesWithTrnID(myInvoiceTrnID, myWindowTitle, myTable, myRefersTo, myOppositeTable)

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
    Dim tmpRecordset As Recordset
    
    'Local μεταβλητές
    Dim lngRow As Long
    Dim lngRowsToAdd  As Long
    Dim bytLoop As Byte
    Dim tmpTableData As typTableData
        
    'Αρχικές τιμές
    lngRow = 0
    lblTitle.Caption = myWindowTitle
    txtTable.text = myTable
    txtRefersTo.text = myRefersTo
    txtOppositeTable.text = myOppositeTable
    CustomizeGrid grdPersonsTransactions
    If grdPersonsTransactions.RowCount = 0 Then AddGridLines grdPersonsTransactions, txtRefersTo.text, 9
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Κύριο SQL
    strSQL = "SELECT InvoiceIssueDate, InvoiceNo, InvoiceCodeID, InvoiceRemarks, InvoiceGrossAmount, InvoiceTrnID, InvoicePersonID, " & txtTable.text & ".Description AS Person, InvoiceInDate, InvoiceInTime " _
        & "FROM Invoices " _
        & "INNER JOIN " & txtTable.text & " ON Invoices.InvoicePersonID = " & txtTable.text & ".ID "
        
    'TrnID παραστατικού
    strThisParameter = "lngInvoiceID Long"
    strThisQuery = "Invoices.InvoiceTrnID = lngInvoiceID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(myInvoiceTrnID)
        
    'Ταξινόμηση
    strOrder = " ORDER BY InvoiceID"
        
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
    If rstRecordset.RecordCount = 0 Then blnError = False: FindInvoicesWithTrnID = True: Exit Function
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        Do While Not .EOF
            txtInvoiceTrnID.text = !InvoiceTrnID
            mskInvoiceIssueDate.text = format(!InvoiceIssueDate, "dd/mm/yyyy")
            txtInvoiceCodeID.text = !InvoiceCodeID
            txtInvoiceNo.text = !InvoiceNo
            txtInvoiceRemarks.text = IIf(IsNull(!InvoiceRemarks), "", !InvoiceRemarks)
            txtInvoiceInDate.text = format(!InvoiceInDate, "dd/mm/yy")
            txtInvoiceInTime.text = format(!InvoiceInTime, "hh:mm")
            'Παραστατικό
            Set tmpRecordset = CheckForMatch("CommonDB", txtInvoiceCodeID.text, "Codes", "CodeID", "Numeric", 0, 1)
            txtInvoiceCodeID.text = tmpRecordset.Fields(0)
            txtCodeShortDescription.text = tmpRecordset.Fields(1)
            lblCodeDescription.Caption = tmpRecordset.Fields(2)
            While Not .EOF
                With grdPersonsTransactions
                    lngRow = lngRow + 1
                    .CellValue(lngRow, "ID") = rstRecordset!InvoicePersonID
                    .CellValue(lngRow, "Description") = rstRecordset!Person
                    .CellValue(lngRow, "Amount") = rstRecordset!InvoiceGrossAmount
                End With
                .MoveNext
            Wend
        Loop
    End With
    
    'Τελικές ενέργειες
    CustomizeGrid grdPersonsTransactions
    EnableFields mskInvoiceIssueDate, txtCodeShortDescription, txtInvoiceNo, txtInvoiceRemarks, grdPersonsTransactions, grdPersonsTransactionsChecks, cmdIndex(0), cmdIndex(1), btnPanel(1)
    EnableGrid grdPersonsTransactions, True
    ColorizeGrid grdPersonsTransactions, grdPersonsTransactionsChecks
    EnableTabStop grdPersonsTransactions, grdPersonsTransactionsChecks
    mskTotal.text = format(CalculateColumnTotal(grdPersonsTransactions, "Amount"), "#,##0.00")
    UpdateButtons Me, 6, 0, 1, 1, 0, 0, 1, 0
    
    FindInvoicesWithTrnID = True
    
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
    FindInvoicesWithTrnID = False
    DisplayErrorMessage True, Err.Description

End Function

Private Function ShowReport()

    With CommonTransactionsIndex
        .lblTitle.Caption = WindowTitle(lblTitle.Caption)
        .txtTable.text = txtTable.text
        .txtOppositeTable.text = txtOppositeTable.text
        .txtRefersTo.text = txtRefersTo.text
        .Tag = "True"
        .Show 1, Me
    End With

End Function

Private Sub btnPanel_Click(Index As Integer)

    Dim intLoop As Integer
    
    For intLoop = 0 To 1
        btnPanel(intLoop).Enabled = True
        frmFrame(intLoop).Visible = False
        shpBridge(intLoop).Visible = False
    Next intLoop
    
    btnPanel(Index).Enabled = False
    frmFrame(Index).Visible = True
    shpBridge(Index).Visible = True
    
    Select Case Index
        'Στοιχεία
        Case 0
            If cmdButton(1).Enabled Then
                If mskInvoiceIssueDate.Enabled Then
                    mskInvoiceIssueDate.SetFocus
                    grdPersonsTransactions.CurCol = 0
                    grdPersonsTransactions.EnsureVisibleRow 1
                End If
            End If
        'Αξιόγραφα
        Case 1
            If cmdButton(1).Enabled Then
                If grdPersonsTransactionsChecks.Enabled And grdPersonsTransactionsChecks.RowCount > 0 Then
                    With grdPersonsTransactionsChecks
                        .SetCurCell 1, 2
                        .SetFocus
                        .TabStop = True
                    End With
                End If
            End If
    End Select

End Sub

Private Sub cmdButton_Click(Index As Integer)
                                                                                                                                
    Select Case Index
        Case 0
            NewRecord
        Case 1
            SaveRecord
        Case 2
            DeleteRecord
        Case 3
            ShowReport
        Case 4
            ShowPersonLedger _
                grdPersonsTransactions.CellText(grdPersonsTransactions.CurRow, "ID"), _
                grdPersonsTransactions.CellText(grdPersonsTransactions.CurRow, "Description"), _
                IIf(txtRefersTo.text = "3", "Καρτέλα προμηθευτή", "Καρτέλα πελάτη"), _
                txtTable.text, _
                txtOppositeTable.text, _
                txtRefersTo.text
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
            'Παραστατικό
            If txtCodeShortDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtCodeShortDescription.text, "Codes", "CodeShortDescription", "String", Val(txtRefersTo.text), 1)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 3, 0, 1, 2, "ID", "Συντ.", "Περιγραφή", 0, 5, 50, 1, 1, 0)
            txtInvoiceCodeID.text = tmpTableData.strCode
            txtCodeShortDescription.text = tmpTableData.strOneField
            lblCodeDescription.Caption = tmpTableData.strTwoField
        Case 1
            'Παραστατικό
            With UtilsCodes
                .Tag = "True"
                .txtRefersTo.text = txtRefersTo.text
                .Show 1, Me
            End With
    End Select

End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        Me.Tag = "False"
    End If
    
    'AddDummyLines grdPersonsTransactions, "99999", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "-9.999.999"
    'AddDummyLines grdPersonsTransactionsChecks, 6, 30, 30, 30, 10, 10

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
        Case vbKeyPageUp
            GotoPreviousPanel Me, btnPanel.Count
        Case vbKeyPageDown
            GotoNextPanel Me, btnPanel.Count
        Case vbKeyF12
            ToggleInfoPanel Me
End Select

End Function

Private Function GotoNextPanel(formName, ParamArray panels())

    Dim intLoop As Integer
    
    For intLoop = 0 To btnPanel.Count - 1
    
        If Not btnPanel(intLoop).Enabled Then
            If intLoop + 1 <= btnPanel.Count - 1 Then
                If btnPanel(intLoop + 1).Enabled Then
                    btnPanel_Click intLoop + 1
                    Exit Function
                End If
            End If
        End If
    
    Next intLoop

End Function

Private Function GotoPreviousPanel(formName, intPanelCount)

    Dim intLoop As Integer
    
    For intLoop = 0 To formName.btnPanel.Count - 1
    
        If Not formName.btnPanel(intLoop).Enabled Then
            If intLoop - 1 >= 0 Then
                If formName.btnPanel(intLoop - 1).Enabled Then
                    btnPanel_Click intLoop - 1
                    Exit Function
                End If
            End If
        End If
    
    Next intLoop

End Function

Private Sub Form_Load()
    
    AddColumnsToGrid grdPersonsTransactions, 25, GetSetting(strAppTitle, "Layout Strings", "grdPersonsTransactions"), "04NCNID,40YLNDescription,10YRFAmount", "ID,Επωνυμία,Ποσό"
    AddColumnsToGrid grdPersonsTransactionsChecks, 25, GetSetting(strAppTitle, "Layout Strings", "grdPersonsTransactionsChecks"), "04NLNBankID,40YLNBankDescription,04NCNCheckIssuedByID,40YLNIssuedByDescription,10YCNCheckNo,10YCDCheckExpire,10YRFCheckAmount", "ID,Τράπεζα,ID εκδότη,Εκδότης,Νο,Λήξη,Ποσό"
    SetUpGrid lstIconList, grdPersonsTransactions
    SetUpGrid lstIconList, grdPersonsTransactionsChecks
    PositionPanels
    PositionControls Me, False: ColorizeControls Me, , True
    ClearFields txtInvoiceTrnID, txtInvoiceCodeID, txtInvoiceInDate, txtInvoiceInTime, mskInvoiceIssueDate, txtCodeShortDescription, lblCodeDescription, txtInvoiceNo, txtInvoiceRemarks, grdPersonsTransactions, grdPersonsTransactionsChecks, mskTotal, mskTotalChecks
    DisableFields mskInvoiceIssueDate, txtCodeShortDescription, txtInvoiceNo, txtInvoiceRemarks
    DisableFields grdPersonsTransactions, grdPersonsTransactionsChecks
    DisableFields cmdIndex(0), cmdIndex(1)
    DisableFields btnPanel(0), btnPanel(1)
    UpdateButtons Me, 6, 1, 0, 0, 1, 0, 0, 1
    ColorizeGrid grdPersonsTransactions, grdPersonsTransactionsChecks

End Sub

Private Sub grdPersonsTransactionsChecks_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case lCol
        Case 2
            'Τράπεζα
            If grdPersonsTransactionsChecks.CellValue(lRow, 2) <> "" Then
                Set tmpRecordset = CheckForMatch("CommonDB", grdPersonsTransactionsChecks.CellValue(lRow, lCol), "Banks", "BankDescription", "String", 1, 1)
                tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                grdPersonsTransactionsChecks.CellValue(lRow, "BankID") = tmpTableData.strCode
                grdPersonsTransactionsChecks.CellValue(lRow, "BankDescription") = tmpTableData.strOneField
                If tmpTableData.strCode <> "" Then
                    MoveToNextColumn grdPersonsTransactionsChecks, lRow, lCol
                End If
            Else
                FillCellWithSomething grdPersonsTransactionsChecks, "", grdPersonsTransactionsChecks.CurRow, "1"
            End If
        Case 4
            'Συναλλασόμενος
            If grdPersonsTransactionsChecks.CellValue(lRow, 4) <> "" Then
                Set tmpRecordset = CheckForMatch("CommonDB", grdPersonsTransactionsChecks.CellValue(lRow, lCol), txtOppositeTable.text, "Description", "String", 1, 1)
                tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 3, 0, 1, 2, "ID", "Περιγραφή", "Α.Φ.Μ.", 0, 50, 10, 1, 0, 1)
                grdPersonsTransactionsChecks.CellValue(lRow, "CheckIssuedByID") = tmpTableData.strCode
                grdPersonsTransactionsChecks.CellValue(lRow, "IssuedByDescription") = tmpTableData.strOneField
                If tmpTableData.strCode <> "" Then
                    MoveToNextColumn grdPersonsTransactionsChecks, lRow, lCol
                End If
            Else
                FillCellWithSomething grdPersonsTransactionsChecks, "", grdPersonsTransactionsChecks.CurRow, "3"
            End If
        Case 5
            'Νο αξιογράφου - Αν είμαι σε κίνηση προμηθευτή, κοιτάζω αν η επιταγή είναι από πελάτη
            If grdPersonsTransactionsChecks.CellValue(lRow, 5) <> "" And txtRefersTo.text = "3" Then
                Set tmpRecordset = NewCheckForMatch("CommonDB", "BankID, BankDescription, ID, Description, CheckNo, CheckExpireDate, CheckAmount", "((Checks", _
                    "INNER JOIN Banks ON Checks.CheckBankID = Banks.BankID) " & _
                    "INNER JOIN Invoices ON Checks.CheckTrnID = Invoices.InvoiceTrnID) " & _
                    "INNER JOIN " & txtOppositeTable.text & " ON Invoices.InvoicePersonID = " & txtOppositeTable.text & ".ID", _
                    "CheckRefersToID = " & Val(txtRefersTo.text) + 1 & " AND InStr(CheckNo,'" & grdPersonsTransactionsChecks.CellValue(lRow, lCol) & "')", "", "")
                If tmpRecordset.RecordCount >= 1 Then
                    tmpTableData = DisplayIndex(tmpRecordset, True, True, "Ευρετήριο", 7, 0, 1, 2, 3, 4, 5, 6, "BankID", "Τράπεζα", "IssuedByID", "Επωνυμία πελάτη", "Νο αξιογράφου", "Λήξη", "Ποσό", 0, 40, 0, 50, 15, 10, 10, 1, 0, 1, 0, 0, 1, 2)
                    If tmpTableData.strCode <> "" Then
                        grdPersonsTransactionsChecks.CellValue(lRow, "BankID") = tmpTableData.strCode
                        grdPersonsTransactionsChecks.CellValue(lRow, "BankDescription") = tmpTableData.strOneField
                        grdPersonsTransactionsChecks.CellValue(lRow, "CheckIssuedByID") = tmpTableData.strTwoField
                        grdPersonsTransactionsChecks.CellValue(lRow, "IssuedByDescription") = tmpTableData.strThreeField
                        grdPersonsTransactionsChecks.CellValue(lRow, "CheckNo") = tmpTableData.strFourField
                        grdPersonsTransactionsChecks.CellValue(lRow, "CheckExpire") = tmpTableData.strFiveField
                        grdPersonsTransactionsChecks.CellValue(lRow, "CheckAmount") = tmpTableData.strSixField
                        mskTotalChecks.text = format(CalculateColumnTotal(grdPersonsTransactionsChecks, "CheckAmount"), "#,##0.00")
                    End If
                End If
            End If
            If grdPersonsTransactionsChecks.CellValue(lRow, 5) <> "" Then
                MoveToNextColumn grdPersonsTransactionsChecks, lRow, lCol
            End If
        Case 6
            'Ημερομηνία λήξης
            If grdPersonsTransactionsChecks.CellValue(lRow, 6) <> "" Then MoveToNextColumn grdPersonsTransactionsChecks, lRow, lCol
        Case 7
            'Ποσό
            mskTotalChecks.text = format(CalculateColumnTotal(grdPersonsTransactionsChecks, "CheckAmount"), "#,##0.00")
            If grdPersonsTransactionsChecks.CellText(lRow, "CheckAmount") <> "" Then
                MoveToNextColumn grdPersonsTransactionsChecks, lRow, lCol
            End If
    End Select
    
    blnGridEditInProgress = False

End Sub

Private Sub grdPersonsTransactionsChecks_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdPersonsTransactionsChecks_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    Dim CtrlDown
    
    CtrlDown = Shift + vbCtrlMask

    'Διαγραφή γραμμής CTRL + DEL
    If KeyCode = 46 And CtrlDown = 4 Then
        FillCellWithSomething grdPersonsTransactionsChecks, "", grdPersonsTransactionsChecks.CurRow, "1,2,3,4,5,6,7"
        mskTotalChecks.text = format(CalculateColumnTotal(grdPersonsTransactionsChecks, "CheckAmount"), "#,##0.00")
    End If

End Sub

Private Sub grdPersonsTransactionsChecks_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid300_10Tec.ETextEditFlags)

    blnGridEditInProgress = True

    If lCol = 7 Then
        If CheckForAcceptableKey(iKeyAscii) Then
            CaptureNumbers grdPersonsTransactionsChecks.TextEditText, lRow, lCol, iKeyAscii, True
        Else
            bCancel = True
        End If
    End If

End Sub

Private Sub grdPersonsTransactionsChecks_TextEditKeyPress(ByVal lRow As Long, ByVal lCol As Long, KeyAscii As Integer)

    If lCol = 7 Then
        If CheckForAcceptableKey(KeyAscii) Then
            CaptureNumbers grdPersonsTransactionsChecks.TextEditText, lRow, lCol, KeyAscii, True
        Else
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub grdPersonsTransactions_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case lCol
        Case 2
            'Επωνυμία
            If grdPersonsTransactions.CellValue(lRow, 2) <> "" Then
                Set tmpRecordset = CheckForMatch("CommonDB", grdPersonsTransactions.CellValue(lRow, lCol), txtTable.text, "Description", "String", 1, 2)
                tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 4, 0, 1, 2, 13, "ID", "Περιγραφή", "Α.Φ.Μ.", "Ε", 0, 50, 15, 0, 1, 0, 1, 1, "Persons")
                grdPersonsTransactions.CellValue(lRow, 1) = tmpTableData.strCode
                grdPersonsTransactions.CellValue(lRow, 2) = tmpTableData.strOneField
                If tmpTableData.strCode <> "" Then
                    cmdButton(4).Enabled = IIf(CheckForLoadedForm("PersonsLedger"), ChangeEditButtonStatus(grdPersonsTransactions, lRow, "ID"), 0)
                    MoveToNextColumn grdPersonsTransactions, lRow, lCol
                End If
            Else
                FillCellWithSomething grdPersonsTransactions, "", grdPersonsTransactions.CurRow, "1"
                mskTotal.text = format(CalculateColumnTotal(grdPersonsTransactions, "Amount"), "#,##0.00")
                cmdButton(4).Enabled = ChangeEditButtonStatus(grdPersonsTransactions, lRow, "ID")
            End If
        Case 3
            'Ποσό
            mskTotal.text = format(CalculateColumnTotal(grdPersonsTransactions, "Amount"), "#,##0.00")
            If grdPersonsTransactions.CellText(lRow, "Amount") <> "" Then
                MoveToNextColumn grdPersonsTransactions, lRow, lCol
            End If
    End Select
    
    blnGridEditInProgress = False
    
End Sub

Private Sub grdPersonsTransactions_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid300_10Tec.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)

    If lCol = 3 Then vNewValue = CheckForMaxLength(sNewText, 12, "Float")

End Sub

Private Sub grdPersonsTransactions_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)

    cmdButton(4).Enabled = Not CheckForLoadedForm("PersonsLedger")
    
    If cmdButton(4).Enabled Then cmdButton(4).Enabled = CheckToEnableButton(grdPersonsTransactions, lRow, "ID")
    
End Sub

Private Sub grdPersonsTransactions_GotFocus()

    If Not grdPersonsTransactions.Enabled Then Exit Sub
    
    Select Case strGridFocus
        Case Is = "FromTop"
            grdPersonsTransactions.SetCurCell 1, 2
            strGridFocus = ""
        Case Is = "FromBottom"
            grdPersonsTransactions.SetCurCell grdPersonsTransactions.RowCount, 2
            strGridFocus = ""
    End Select
    
    cmdButton(4).Enabled = Not CheckForLoadedForm("PersonsLedger")
    
    If cmdButton(4).Enabled Then cmdButton(4).Enabled = CheckToEnableButton(grdPersonsTransactions, grdPersonsTransactions.CurRow, "ID")
    
End Sub

Private Sub grdPersonsTransactions_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdPersonsTransactions_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    Dim CtrlDown
    
    CtrlDown = Shift + vbCtrlMask
    
    'F5 Πίνακας
    If KeyCode = vbKeyF5 Then
        With Persons
            .txtTable.text = txtTable.text
            .txtRefersTo.text = txtRefersTo.text
            .Tag = "True"
            .Show 1, Me
        End With
    End If

    'Πάνω βελάκι
    If KeyCode = 38 Then
        If grdPersonsTransactions.CurRow = 1 Then
            grdPersonsTransactions.CurCol = 0
            txtInvoiceRemarks.SetFocus
            Exit Sub
        End If
    End If
    
    'Κάτω βελάκι
    If KeyCode = 40 Then
        If grdPersonsTransactions.CurRow = grdPersonsTransactions.RowCount Then
            grdPersonsTransactions.CurCol = 0
            mskInvoiceIssueDate.SetFocus
            Exit Sub
        End If
    End If
    
    'Διαγραφή γραμμής CTRL + DEL
    If KeyCode = 46 And CtrlDown = 4 Then
        FillCellWithSomething grdPersonsTransactions, "", grdPersonsTransactions.CurRow, "1,2,3"
        mskTotal.text = format(CalculateColumnTotal(grdPersonsTransactions, "Amount"), "#,##0.00")
    End If
    
End Sub

Private Sub grdPersonsTransactions_LostFocus()

    cmdButton(4).Enabled = False
    
End Sub

Private Sub grdPersonsTransactions_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid300_10Tec.ETextEditFlags)

    blnGridEditInProgress = True
    
    If lCol = 1 Or (lCol = 3 And grdPersonsTransactions.CellValue(lRow, 1) = "") Then bCancel = True
    
    If lCol = 3 Then
        If CheckForAcceptableKey(iKeyAscii) Then
            CaptureNumbers grdPersonsTransactions.TextEditText, lRow, lCol, iKeyAscii, True
        Else
            bCancel = True
        End If
    End If
        
End Sub

Private Sub grdPersonsTransactions_TextEditKeyPress(ByVal lRow As Long, ByVal lCol As Long, KeyAscii As Integer)

    If lCol = 3 Then
        If CheckForAcceptableKey(KeyAscii) Then
            CaptureNumbers grdPersonsTransactions.TextEditText, lRow, lCol, KeyAscii, True
        Else
            KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdPersonsTransactions", grdPersonsTransactions.LayoutCol
    SaveSetting strAppTitle, "Layout Strings", "grdPersonsTransactionsChecks", grdPersonsTransactionsChecks.LayoutCol

End Sub

Private Function ValidateFields()

    Dim lngRow As Long
    Dim lngCol As Long
    Dim blnSomethingIsGiven As Boolean
    Dim blnLineIsCorrect As Boolean
    
    ValidateFields = False
    
    'Ημερομηνία
    If Not CheckDateWithinLimits(strAppTitle, mskInvoiceIssueDate.text, datClosedPeriod) Then
        btnPanel_Click 0
        mskInvoiceIssueDate.SetFocus
        Exit Function
    End If
    
    'Τύπος παραστατικού
    If DisplayMessage(1, 4, 1, "", txtInvoiceCodeID.text) Then
        btnPanel_Click 0
        txtCodeShortDescription.SetFocus
        Exit Function
    End If
    
    'Νο παραστατικού
    If DisplayMessage(1, 4, 1, "", txtInvoiceNo.text) Then
        btnPanel_Click 0
        txtInvoiceNo.SetFocus
        Exit Function
    End If
    
    'Νο παραστατικού = ακέραιος
    If Not CheckForInteger(txtInvoiceNo.text) Then
        If DisplayMessage(2, 4, 1, "", "") Then
            btnPanel_Click 0
            txtInvoiceNo.SetFocus
            Exit Function
        End If
    End If
    
    'Συναλλασόμενοι
    With grdPersonsTransactions
        blnLineIsCorrect = False
        For lngRow = 1 To .RowCount
            blnSomethingIsGiven = False
            For lngCol = 1 To .colCount
                If .CellText(lngRow, lngCol) <> "" Then
                    blnSomethingIsGiven = True
                End If
            Next lngCol
            If blnSomethingIsGiven Then
                If .CellValue(lngRow, "ID") = "" Or .CellValue(lngRow, "Amount") = "" Then
                    btnPanel_Click 0
                    If MyMsgBox(4, strAppTitle, strMessages(11) & lngRow & " δεν είναι σωστή.", 1) Then
                    End If
                    .SetCurCell lngRow, "Amount"
                    .SetFocus
                    Exit Function
                Else
                    blnLineIsCorrect = True
                End If
            End If
        Next lngRow
        If Not blnLineIsCorrect Then
            btnPanel_Click 0
            If DisplayMessage(9, 4, 1, "", "") Then
            End If
            txtInvoiceRemarks.SetFocus
            Exit Function
        End If
    End With
    
    'Αξιόγραφα
    With grdPersonsTransactionsChecks
        For lngRow = 1 To .RowCount
            blnSomethingIsGiven = False
            For lngCol = 1 To .colCount
                If .CellText(lngRow, lngCol) <> "" Then
                    blnSomethingIsGiven = True
                End If
            Next lngCol
            If blnSomethingIsGiven Then
                If .CellText(lngRow, "BankID") = "" Or .CellText(lngRow, "CheckNo") = "" Or .CellText(lngRow, "CheckExpire") = "" Or .CellText(lngRow, "CheckAmount") = "" Then
                    btnPanel_Click 1
                    If MyMsgBox(4, strAppTitle, strMessages(11) & lngRow & " δεν είναι σωστή.", 1) Then
                    End If
                    .SetCurCell lngRow, "BankDescription"
                    .SetFocus
                    Exit Function
                End If
            End If
        Next lngRow
    End With
    
    ValidateFields = True

End Function

Private Sub mskInvoiceIssueDate_GotFocus()

    strGridFocus = "FromBottom"

End Sub

Private Sub txtCodeShortDescription_Change()

    If txtCodeShortDescription.text = "" Then ClearFields txtInvoiceCodeID, lblCodeDescription

End Sub

Private Sub txtCodeShortDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0
    If KeyCode = vbKeyF5 Then cmdIndex_Click 1

End Sub

Private Sub txtCodeShortDescription_Validate(Cancel As Boolean)

    If txtInvoiceCodeID.text = "" And txtCodeShortDescription.text <> "" Then cmdIndex_Click 0: If txtInvoiceCodeID.text = "" Then Cancel = True

End Sub

Private Sub txtInvoiceRemarks_GotFocus()

    strGridFocus = "FromTop"
    
End Sub

Private Function PositionPanels()

    Dim intLoop As Integer
    
    For intLoop = 0 To 1
        frmFrame(intLoop).Visible = False
    Next intLoop
        
    For intLoop = 0 To 1
        btnPanel(intLoop).Enabled = True
        shpBridge(intLoop).Visible = False
        With frmFrame(intLoop)
            .Height = 7140
            .Width = 10215
            .Left = 1875
            .Top = 1125
            .BackColor = &HE0E0E0
        End With
    Next intLoop
    
    btnPanel(0).Enabled = False
    frmFrame(0).Visible = True
    shpBridge(0).Visible = True

End Function


