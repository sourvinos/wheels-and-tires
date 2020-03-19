VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form ItemsTransactions 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
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
   Begin VB.Frame frmInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2940
      Left            =   10050
      TabIndex        =   5
      Top             =   825
      Width           =   4515
      Begin VB.TextBox txtCodeInventoryQty 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3675
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "13"
         Top             =   1950
         Width           =   780
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   75
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Codes.CodeInventoryQty"
         Top             =   1950
         Width           =   3540
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
         Left            =   3675
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "4"
         Top             =   1200
         Width           =   780
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
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceInTime"
         Top             =   1200
         Width           =   3540
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
         Left            =   3675
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "3"
         Top             =   825
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
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceInDate"
         Top             =   825
         Width           =   3540
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
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "RefersTo"
         Top             =   1575
         Width           =   3540
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
         Left            =   3675
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "6"
         Top             =   1575
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
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceCodeID"
         Top             =   450
         Width           =   3540
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
         Left            =   3675
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "2"
         Top             =   450
         Width           =   780
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
         Left            =   3675
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "1"
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
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceTrnID"
         Top             =   75
         Width           =   3540
      End
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   75
         Top             =   2325
         _ExtentX        =   953
         _ExtentY        =   953
         IconSizeX       =   26
         IconSizeY       =   32
         Size            =   14064
         Images          =   "ItemsTransactions.frx":0000
         Version         =   131072
         KeyCount        =   4
         Keys            =   "ÿÿÿ"
      End
   End
   Begin UserControls.newInteger mskTotal 
      Height          =   465
      Left            =   10200
      TabIndex        =   33
      Top             =   6900
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   820
      Alignment       =   1
      ForeColor       =   0
      Text            =   "0"
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
   Begin iGrid300_10Tec.iGrid grdItemsTransactions 
      Height          =   3615
      Left            =   2175
      TabIndex        =   32
      Top             =   3225
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   6376
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
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   450
      TabIndex        =   16
      Top             =   7875
      Width           =   10365
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   17
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
         TabIndex        =   18
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
         TabIndex        =   19
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
         TabIndex        =   20
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
         TabIndex        =   21
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
         TabIndex        =   22
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
         TabIndex        =   23
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
   Begin UserControls.newDate mskInvoiceIssueDate 
      Height          =   465
      Left            =   2175
      TabIndex        =   1
      Top             =   1125
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
   Begin UserControls.newText txtInvoiceRemarks 
      Height          =   465
      Left            =   2175
      TabIndex        =   4
      Top             =   2700
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
      Top             =   1650
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
      TabIndex        =   24
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
      PicNormal       =   "ItemsTransactions.frx":3710
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   1
      Left            =   4200
      TabIndex        =   25
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
      PicNormal       =   "ItemsTransactions.frx":3CAA
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newText txtInvoiceNo 
      Height          =   465
      Left            =   2175
      TabIndex        =   3
      Top             =   2175
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
      TabIndex        =   31
      Top             =   1725
      Width           =   4365
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
      Left            =   9225
      TabIndex        =   30
      Top             =   6975
      Width           =   540
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
      TabIndex        =   29
      Top             =   1725
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
      TabIndex        =   28
      Top             =   2250
      Width           =   1290
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
      TabIndex        =   27
      Top             =   2775
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
      Top             =   1575
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
      TabIndex        =   26
      Top             =   1200
      Width           =   1290
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   4
      Left            =   9750
      Top             =   6825
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   3750
      Top             =   8550
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   10425
      Top             =   7350
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   11625
      Top             =   4650
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
      Caption         =   "Κινήσεις ειδών"
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
      TabIndex        =   0
      Top             =   75
      Width           =   3480
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
      Left            =   2850
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
Attribute VB_Name = "ItemsTransactions"
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
        grdItemsTransactions.CancelEdit
        Exit Function
    End If
    
    If Not blnStatus Then
        If MyMsgBox(3, strAppTitle, strMessages(3), 2) Then
            blnStatus = False
            ClearFields txtInvoiceTrnID, txtInvoiceCodeID, txtInvoiceInDate, txtInvoiceInTime, mskInvoiceIssueDate, txtCodeShortDescription, lblCodeDescription, txtInvoiceNo, txtInvoiceRemarks, grdItemsTransactions, mskTotal
            DisableFields mskInvoiceIssueDate, txtCodeShortDescription, txtInvoiceNo, txtInvoiceRemarks, grdItemsTransactions, cmdIndex(0), cmdIndex(1)
            UpdateButtons Me, 6, 1, 0, 0, IIf(CheckForLoadedForm("CommonTransactionsIndex"), 0, 1), 0, 0, 1
        End If
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Function DeleteInvoicesTrn()

    On Error GoTo ErrTrap
    
    Dim strSQL As String
    
    If blnError Then Exit Function
        
    strSQL = "DELETE FROM InvoicesTrn WHERE InvoiceTrnID = " & Val(txtInvoiceTrnID.text)
    CommonDB.Execute (strSQL)
    
    Exit Function
    
ErrTrap:
    blnError = True
    DeleteInvoicesTrn = False
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
    DeleteInvoicesTrn
    
    If Not blnError Then
        CommitTrans
        ClearFields txtInvoiceTrnID, txtInvoiceCodeID, txtInvoiceInDate, txtInvoiceInTime, mskInvoiceIssueDate, txtCodeShortDescription, lblCodeDescription, txtInvoiceNo, txtInvoiceRemarks, grdItemsTransactions, mskTotal
        DisableFields mskInvoiceIssueDate, txtCodeShortDescription, txtInvoiceNo, txtInvoiceRemarks, grdItemsTransactions, cmdIndex(0), cmdIndex(1)
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

Function DoSharedStuff(myInvoiceTrnID, myWindowTitle)

    FindInvoicesWithTrnID myInvoiceTrnID, myWindowTitle
    FindItemsWithTrnID myInvoiceTrnID
    Me.Tag = "False"
    If Me.Visible Then
        Unload CommonTransactionsIndex
        Me.mskInvoiceIssueDate.SetFocus
    Else
        Me.Show 1
    End If

End Function

Function FindItemsWithTrnID(myInvoiceTrnID)

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
    
    'Qty
    Dim intThisQty As Integer
    Dim intLastQty As Integer
        
    'Αρχικές τιμές
    lngRow = 0
    
    If grdItemsTransactions.RowCount = 0 Then AddGridLines grdItemsTransactions, txtRefersTo.text, 100
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Κύριο SQL
    strSQL = "SELECT InvoicesTrn.ItemID, Qty, ItemDescription, CategoryShortDescription, CategoryID, CategoryDescription, ManufacturerID, ManufacturerDescription, ItemBalance " _
        & "FROM ((InvoicesTrn " _
        & "INNER JOIN Items ON InvoicesTrn.ItemID = Items.ItemID) " _
        & "INNER JOIN Categories ON Items.ItemCategoryID = Categories.CategoryID) " _
        & "INNER JOIN Manufacturers ON Items.ItemManufacturerID = Manufacturers.ManufacturerID "

    'TrnID ειδών
    strThisParameter = "lngInvoiceTrnID Long"
    strThisQuery = "InvoiceTrnID = lngInvoiceTrnID "
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(myInvoiceTrnID)
        
    'Ταξινόμηση
    strOrder = " ORDER BY InvoicesTrn.ID"
        
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
    If rstRecordset.RecordCount = 0 Then blnError = False: FindItemsWithTrnID = True: Exit Function
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        While Not .EOF
            With grdItemsTransactions
                lngRow = lngRow + 1
                grdItemsTransactions.CellValue(lngRow, "ItemID") = rstRecordset!ItemID
                grdItemsTransactions.CellValue(lngRow, "ItemDescription") = rstRecordset!ItemDescription
                grdItemsTransactions.CellValue(lngRow, "CategoryID") = rstRecordset!CategoryID
                grdItemsTransactions.CellValue(lngRow, "CategoryShortDescription") = rstRecordset!CategoryShortDescription
                grdItemsTransactions.CellValue(lngRow, "CategoryDescription") = rstRecordset!CategoryDescription
                grdItemsTransactions.CellValue(lngRow, "ManufacturerID") = rstRecordset!ManufacturerID
                grdItemsTransactions.CellValue(lngRow, "ManufacturerDescription") = rstRecordset!ManufacturerDescription
                grdItemsTransactions.CellValue(lngRow, "Qty") = rstRecordset!Qty
                
                
                intThisQty = .CellValue(lngRow, "Qty")
                
                '
                If txtCodeInventoryQty.text = "+" Then
                    intLastQty = rstRecordset!ItemBalance - intThisQty
                End If
                If txtCodeInventoryQty.text = "-" Then
                    intLastQty = rstRecordset!ItemBalance + intThisQty
                End If
                If txtCodeInventoryQty.text = "" Then
                    intLastQty = rstRecordset!ItemBalance
                End If
                '
                
                .CellValue(lngRow, "LastQty") = intLastQty
                
            End With
            .MoveNext
        Wend
    End With
    
    'Τελικές ενέργειες
    mskTotal.text = format(CalculateColumnTotal(grdItemsTransactions, "Qty"), "#,##0")
    EnableGrid grdItemsTransactions, True
    FindItemsWithTrnID = True
    
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
    FindItemsWithTrnID = False
    DisplayErrorMessage True, Err.Description

End Function

Private Function NewRecord()
    
    blnStatus = True
    ClearFields txtInvoiceTrnID, txtInvoiceCodeID, txtInvoiceInDate, txtInvoiceInTime, mskInvoiceIssueDate, txtCodeShortDescription, lblCodeDescription, txtInvoiceNo, txtInvoiceRemarks, grdItemsTransactions, mskTotal
    EnableFields mskInvoiceIssueDate, txtCodeShortDescription, txtInvoiceNo, txtInvoiceRemarks, grdItemsTransactions, cmdIndex(0), cmdIndex(1)
    EditableFields grdItemsTransactions
    CustomizeGrid grdItemsTransactions
    EnableTabStop grdItemsTransactions
    AddGridLines grdItemsTransactions, txtRefersTo.text, 100
    InitializeFields mskInvoiceIssueDate, mskTotal
    UpdateButtons Me, 6, 0, 1, 0, 0, 0, 1, 0
    mskInvoiceIssueDate.SetFocus

End Function

Private Function SaveInvoicesTrn()

    Dim lngRow As Long
    
    If blnError Then Exit Function
    
    With grdItemsTransactions
        For lngRow = 1 To .RowCount
            If .CellValue(lngRow, "ItemID") <> "" Then
                If Not MainSaveRecord("CommonDB", "InvoicesTrn", True, strAppTitle, "InvoiceTrnID", txtInvoiceTrnID.text, _
                    .CellValue(lngRow, "ItemID"), _
                    .CellValue(lngRow, "Qty"), _
                    0, _
                    0, _
                    0, _
                    0, _
                    "", _
                    0, _
                    0, _
                    0, _
                    0, _
                    lngTrnID) <> 0 Then
                    blnError = True
                End If
            End If
        Next lngRow
    End With
    
    SaveInvoicesTrn = True
    
End Function

Private Function SaveInvoice()

    Dim lngRow As Long
    
    If blnError Then Exit Function
    
    lngTrnID = IIf(txtInvoiceTrnID.text = "", AddOneToTheLastRecord, txtInvoiceTrnID.text)
    
        If Not MainSaveRecord("CommonDB", "Invoices", blnStatus, strAppTitle, "TrnID", _
            txtInvoiceTrnID.text, _
            mskInvoiceIssueDate.text, Val(txtInvoiceNo.text), txtInvoiceCodeID.text, Val(txtRefersTo.text), _
            "0", "0", "0", "0", "0", "0", "0", "0", _
            lngTrnID, _
            txtInvoiceRemarks.text, _
            "", _
            "6", _
            "0", _
            "0", _
            "0", _
            IIf(blnStatus, Date, txtInvoiceInDate.text), _
            IIf(blnStatus, Time, txtInvoiceInTime.text), _
            "", _
            "", _
            "", _
            "", _
            "1", _
            "0", _
            "0", _
            strCurrentUser) <> 0 Then
            blnError = True
        End If
    
End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    blnError = False
    
    BeginTrans
    
    DeleteInvoicesTrn
    SaveInvoice
    SaveInvoicesTrn
    UpdateItemsWithNewBalance
    
    If Not blnError Then
        CommitTrans
        ClearFields txtInvoiceTrnID, txtInvoiceCodeID, txtInvoiceInDate, txtInvoiceInTime, mskInvoiceIssueDate, txtCodeShortDescription, lblCodeDescription, txtInvoiceNo, txtInvoiceRemarks, grdItemsTransactions, mskTotal
        DisableFields mskInvoiceIssueDate, txtCodeShortDescription, txtInvoiceNo, txtInvoiceRemarks, grdItemsTransactions, cmdIndex(0), cmdIndex(1)
        UpdateButtons Me, 6, 1, 0, 0, IIf(CheckForLoadedForm("CommonTransactionsIndex"), 0, 1), 0, 0, 1
    Else
        Rollback
    End If
    
End Function

Private Function UpdateItemsWithNewBalance()

    '1 = CategoryID
    '2 = CategoryShortDescription
    '3 = ItemID
    '4 = ItemDescription
    '5 = ManufacturerDescription
    '6 = Qty
    '7 = UnitPrice
    '8 = TotalNetPreDiscount
    '9 = DiscPercent
    '10 = DiscAmount
    '11 = DiscAllow
    '12 = TotalNetPostDiscount
    '13 = VATPercent
    '14 = VATAmount
    '15 = TotalGross
    '16 = LastQty
    
    Dim intQty As Integer
    Dim lngRow As Long
    Dim lngItemID As Long
    Dim intLastQty As Integer
    Dim intThisQty As Integer
    Dim intNewQty As Integer
    
    Dim rsItems As Recordset
    
    If blnError Then Exit Function
    
    Set rsItems = CommonDB.OpenRecordset("Items")
    
    With grdItemsTransactions
        For lngRow = 1 To .RowCount
            If .CellValue(lngRow, "ItemID") <> "" Then
                
                lngItemID = .CellValue(lngRow, "ItemID")
                intLastQty = .CellValue(lngRow, "LastQty")
                intThisQty = .CellValue(lngRow, "Qty")
                
                If txtCodeInventoryQty.text = "+" Then
                    intNewQty = intLastQty + intThisQty
                End If
                If txtCodeInventoryQty.text = "-" Then
                    intNewQty = intLastQty - intThisQty
                End If
                If txtCodeInventoryQty.text = "" Then
                    intNewQty = intLastQty
                End If
                
                rsItems.Index = "ID"
                rsItems.Seek "=", .CellValue(lngRow, "ItemID")
                
                If Not rsItems.NoMatch Then
                    rsItems.Edit
                    rsItems!ItemBalance = intNewQty
                    rsItems.Update
                End If
            
            End If
        Next lngRow
    End With

    UpdateItemsWithNewBalance = True

End Function


Function FindInvoicesWithTrnID(myInvoiceTrnID, myWindowTitle)

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
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Κύριο SQL
    strSQL = "SELECT InvoiceIssueDate, InvoiceNo, InvoiceCodeID, InvoiceTrnID, InvoiceRemarks, InvoiceInDate, InvoiceInTime, InvoiceRefersToID " _
        & "FROM Invoices "
        
    'TrnID παραστατικού
    strThisParameter = "lngInvoiceID Long"
    strThisQuery = "Invoices.InvoiceTrnID = lngInvoiceID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(myInvoiceTrnID)
        
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
        txtCodeInventoryQty.text = tmpRecordset.Fields(4)
        'Βοηθητικά
        txtRefersTo.text = !InvoiceRefersToID
    End With
    
    'Τελικές ενέργειες
    CustomizeGrid grdItemsTransactions
    EnableFields mskInvoiceIssueDate, txtCodeShortDescription, txtInvoiceNo, txtInvoiceRemarks, grdItemsTransactions, cmdIndex(0), cmdIndex(1)
    EnableTabStop grdItemsTransactions
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

Private Function ShowLedger(myGrid As iGrid, myGridRow As Long)

    With ItemsLedger
        .txtCategoryID.text = myGrid.CellText(myGridRow, "CategoryID")
        .txtCategoryShortDescription.text = myGrid.CellText(myGridRow, "CategoryShortDescription")
        .lblCategoryDescription.Caption = myGrid.CellText(myGridRow, "CategoryDescription")
        .txtManufacturerID.text = myGrid.CellText(myGridRow, "ManufacturerID")
        .txtManufacturerDescription.text = myGrid.CellText(myGridRow, "ManufacturerDescription")
        .txtItemID.text = myGrid.CellText(myGridRow, "ItemID")
        .txtItemDescription.text = myGrid.CellText(myGridRow, "ItemDescription")
        .Tag = "True"
        DisableFields .txtCategoryShortDescription, .txtManufacturerDescription, .txtItemDescription, .cmdIndex(0), .cmdIndex(1), .cmdIndex(2)
        .Show 1, Me
    End With

End Function

Private Function ShowReport()

    With CommonTransactionsIndex
        .lblTitle.Caption = WindowTitle(lblTitle.Caption)
        .txtRefersTo.text = txtRefersTo.text
        .Tag = "True"
        .Show 1, Me
    End With

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
            ShowReport
        Case 4
            ShowLedger grdItemsTransactions, grdItemsTransactions.CurRow
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
            Set tmpRecordset = CheckForMatch("CommonDB", txtCodeShortDescription.text, "Codes", "CodeShortDescription", "String", txtRefersTo.text, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 4, 0, 1, 2, 4, "ID", "Συντ.", "Περιγραφή", "Ποσότητες", 0, 5, 40, 4, 1, 1, 1, 0)
            txtInvoiceCodeID.text = tmpTableData.strCode
            txtCodeShortDescription.text = tmpTableData.strOneField
            lblCodeDescription.Caption = tmpTableData.strTwoField
            txtCodeInventoryQty.text = tmpTableData.strThreeField
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
    
    'AddDummyLines grdItemsTransactions, 6, 6, 3, 40, 50, 6, 20, 10
    'grdItemsTransactions.Enabled = True

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
    
    AddColumnsToGrid grdItemsTransactions, 25, GetSetting(strAppTitle, "Layout Strings", "grdItemsTransactions"), _
        "14NCNItemID,04NCNCategoryID,12YCNCategoryShortDescription,40NLNCategoryDescription,50YLNItemDescription,06NCNManufacturerID,40NLNManufacturerDescription,10YRIQty,10NRIXLastQty", _
        "ID Είδους,ID Κατηγορίας,Κατ,Κατηγορία,Περιγραφή,ID Κατασκευαστή,Κατασκευαστής,Ποσότητα,Τρέχουσα  ποσότητα"
    SetUpGrid lstIconList, grdItemsTransactions
    PositionControls Me, False: ColorizeControls Me, , False
    ClearFields txtInvoiceTrnID, txtInvoiceCodeID, txtInvoiceInDate, txtInvoiceInTime, mskInvoiceIssueDate, txtCodeShortDescription, lblCodeDescription, txtInvoiceNo, txtInvoiceRemarks, grdItemsTransactions, mskTotal
    DisableFields mskInvoiceIssueDate, txtCodeShortDescription, txtInvoiceNo, txtInvoiceRemarks, grdItemsTransactions, cmdIndex(0), cmdIndex(1)
    UpdateButtons Me, 6, 1, 0, 0, 1, 0, 0, 1

End Sub

Private Sub grdItemsTransactions_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)

    Dim strCategoryID As String
    Dim strItemQuickDescription As String
    Dim strItemDescription As String
    
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case lCol
        Case 3
            'Συντ. Κατηγορίας
            Set tmpRecordset = CheckForMatch("CommonDB", grdItemsTransactions.CellValue(lRow, 3), "Categories", "CategoryShortDescription", "String", 1, 1)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 3, 0, 1, 2, "ID", "Συντ.", "Περιγραφή", 0, 4, 40, 1, 1, 0)
            If tmpTableData.strCode = "" Then
                FillCellWithSomething grdItemsTransactions, "", grdItemsTransactions.CurRow, "1,2,3,4,5,6"
            End If
            If tmpTableData.strCode <> "" Then
                grdItemsTransactions.CellValue(lRow, "CategoryID") = tmpTableData.strCode
                grdItemsTransactions.CellValue(lRow, "CategoryShortDescription") = tmpTableData.strOneField
                grdItemsTransactions.CellValue(lRow, "CategoryDescription") = tmpTableData.strTwoField
                cmdButton(4).Enabled = IIf(CheckForLoadedForm("ItemsLedger"), ChangeEditButtonStatus(grdItemsTransactions, lRow, "ItemID"), 0)
                MoveToNextColumn grdItemsTransactions, lRow, lCol
            End If
        Case 5
            'Είδος
            If grdItemsTransactions.CellValue(lRow, "ItemDescription") <> "" Then
                strCategoryID = IIf(grdItemsTransactions.CellValue(lRow, "CategoryID") <> "", "ItemCategoryID = " & grdItemsTransactions.CellValue(lRow, "CategoryID"), "")
                strItemQuickDescription = IIf(grdItemsTransactions.CellValue(lRow, "ItemDescription") <> "", grdItemsTransactions.CellValue(lRow, "ItemDescription"), "'")
                
                If Left(strItemQuickDescription, 1) <> "*" Then
                    strItemDescription = "Left(ItemQuickDescription, " & Len(strItemQuickDescription) & ") = '" & strItemQuickDescription & "'" & IIf(strCategoryID <> "", " AND " & strCategoryID, "")
                Else
                    If Len(strItemQuickDescription) > 1 Then
                        strItemDescription = "InStr(ItemQuickDescription, " & Right(strItemQuickDescription, Len(strItemQuickDescription) - 1) & ")" & IIf(strCategoryID <> "", " And " & strCategoryID, "")
                    Else
                        strItemDescription = strCategoryID
                    End If
                End If
                
                Set tmpRecordset = NewCheckForMatch("CommonDB", "ItemID, ItemCategoryID, ItemManufacturerID, CategoryDescription, ManufacturerDescription, ItemDescription, CategoryShortDescription, ItemVATPercent, ItemBalance, ItemActive ", _
                    "((Items", _
                    "INNER JOIN Categories ON Items.ItemCategoryID = Categories.CategoryID) " & _
                    "INNER JOIN Manufacturers ON Items.ItemManufacturerID = Manufacturers.ManufacturerID) ", strItemDescription, "", "CategoryDescription, ManufacturerDescription, ItemDescription")
                tmpTableData = DisplayIndex(tmpRecordset, True, True, "Ευρετήριο", 10, 0, 1, 2, 3, 5, 4, 6, 7, 8, 9, "ID", "ID Κατηγορίας", "ID Κατασκευαστή", "Κατηγορία", "Περιγραφή", "Κατασκευαστής", "Συντ. κατηγορίας", "Φ.Π.Α.", "Υπόλοιπο", "Ε", 0, 0, 0, 40, 50, 40, 0, 0, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 1, "Items")
                
                'Set tmpRecordset = NewCheckForMatch("CommonDB", "ItemID, ItemCategoryID, ItemManufacturerID, CategoryDescription, ManufacturerDescription, ItemDescription, CategoryShortDescription, ItemVATPercent", _
                    "((Items", _
                    "INNER JOIN Categories ON Items.ItemCategoryID = Categories.CategoryID) " & _
                    "INNER JOIN Manufacturers ON Items.ItemManufacturerID = Manufacturers.ManufacturerID) ", _
                    "Left(ItemQuickDescription, " & Len(strItemQuickDescription) & ") = '" & strItemQuickDescription & "'" & strCategoryID, "CategoryDescription, ManufacturerDescription, ItemDescription")
                'tmpTableData = DisplayIndex(tmpRecordset, True, True, "Ευρετήριο", 8, 0, 1, 2, 3, 5, 4, 6, 7, "ID", "ID Κατηγορίας", "ID Κατασκευαστή", "Κατηγορία", "Περιγραφή", "Κατασκευαστής", "Συντ. κατηγορίας", "Φ.Π.Α.", 0, 0, 0, 40, 50, 40, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0)
                
                If tmpTableData.strCode = "" Then
                    grdItemsTransactions.CellValue(lRow, "ItemID") = ""
                    grdItemsTransactions.CellValue(lRow, "ItemDescription") = ""
                    grdItemsTransactions.CellValue(lRow, "ManufacturerID") = ""
                    grdItemsTransactions.CellValue(lRow, "ManufacturerDescription") = ""
                Else
                    grdItemsTransactions.CellValue(lRow, "ItemID") = tmpTableData.strCode
                    grdItemsTransactions.CellValue(lRow, "ItemDescription") = tmpTableData.strFourField
                    grdItemsTransactions.CellValue(lRow, "CategoryID") = tmpTableData.strOneField
                    grdItemsTransactions.CellValue(lRow, "CategoryDescription") = tmpTableData.strThreeField
                    grdItemsTransactions.CellValue(lRow, "CategoryShortDescription") = tmpTableData.strSixField
                    grdItemsTransactions.CellValue(lRow, "ManufacturerID") = tmpTableData.strTwoField
                    grdItemsTransactions.CellValue(lRow, "ManufacturerDescription") = tmpTableData.strFiveField
                    grdItemsTransactions.CellValue(lRow, "LastQty") = tmpTableData.strEightField
                    cmdButton(4).Enabled = IIf(CheckForLoadedForm("ItemsLedger"), ChangeEditButtonStatus(grdItemsTransactions, lRow, "ItemID"), 0)
                    MoveToNextColumn grdItemsTransactions, lRow, lCol
                End If
                mskTotal.text = format(CalculateColumnTotal(grdItemsTransactions, "Qty"), "#,##0")
                cmdButton(4).Enabled = ChangeEditButtonStatus(grdItemsTransactions, lRow, "ItemID")
            End If
        Case 8
            'Ποσότητα
            mskTotal.text = format(CalculateColumnTotal(grdItemsTransactions, "Qty"), "#,##0")
            If grdItemsTransactions.CellText(lRow, "Qty") <> "" Then
                MoveToNextColumn grdItemsTransactions, lRow, lCol
            End If
    End Select
    
    blnGridEditInProgress = False
    
End Sub

Private Sub grdItemsTransactions_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid300_10Tec.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)

    If lCol = 6 Then vNewValue = CheckForMaxLength(sNewText, 4, "Integer")

End Sub

Private Sub grdItemsTransactions_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)

    cmdButton(4).Enabled = IIf(Not CheckForLoadedForm("ItemsLedger"), CheckToEnableButton(grdItemsTransactions, lRow, "ItemID"), 0)
    
End Sub

Private Sub grdItemsTransactions_GotFocus()

    If Not grdItemsTransactions.Enabled Then Exit Sub
    
    Select Case strGridFocus
        Case Is = "FromTop"
            grdItemsTransactions.SetCurCell 1, 3
            strGridFocus = ""
        Case Is = "FromBottom"
            grdItemsTransactions.SetCurCell grdItemsTransactions.RowCount, 3
            strGridFocus = ""
    End Select
    
    cmdButton(4).Enabled = IIf(Not CheckForLoadedForm("ItemsLedger"), CheckToEnableButton(grdItemsTransactions, grdItemsTransactions.CurRow, "ItemID"), 0)
    
End Sub

Private Sub grdItemsTransactions_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdItemsTransactions_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    'Πάνω βελάκι
    If KeyCode = 38 Then
        If grdItemsTransactions.CurRow = 1 Then
            grdItemsTransactions.CurCol = 0
            txtInvoiceRemarks.SetFocus
            Exit Sub
        End If
    End If
    
    'Κάτω βελάκι
    If KeyCode = 40 Then
        If grdItemsTransactions.CurRow = grdItemsTransactions.RowCount Then
            grdItemsTransactions.CurCol = 0
            mskInvoiceIssueDate.SetFocus
            Exit Sub
        End If
    End If

End Sub

Private Sub grdItemsTransactions_LostFocus()

    cmdButton(4).Enabled = False
    
End Sub

Private Sub grdItemsTransactions_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid300_10Tec.ETextEditFlags)

    blnGridEditInProgress = True
    
    If lCol <> 3 And lCol <> 5 And lCol <> 8 Then bCancel = True
    
    If lCol = 8 Then
        If CheckForAcceptableKey(iKeyAscii) Then
            CaptureNumbers grdItemsTransactions.TextEditText, lRow, lCol, iKeyAscii, False
        Else
            bCancel = True
        End If
    End If
        
End Sub

Private Sub grdItemsTransactions_TextEditKeyPress(ByVal lRow As Long, ByVal lCol As Long, KeyAscii As Integer)

    If lCol = 8 Then
        If CheckForAcceptableKey(KeyAscii) Then
            CaptureNumbers grdItemsTransactions.TextEditText, lRow, lCol, KeyAscii, False
        Else
            KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdItemsTransactions", grdItemsTransactions.LayoutCol

End Sub

Private Function ValidateFields()

    Dim lngRow As Long
    Dim lngCol As Long
    Dim blnSomethingIsGiven As Boolean
    Dim blnLineIsCorrect As Boolean
    
    ValidateFields = False
    
    'Ημερομηνία
    If Not CheckDateWithinLimits(strAppTitle, mskInvoiceIssueDate.text, datClosedPeriod) Then
        mskInvoiceIssueDate.SetFocus
        Exit Function
    End If
    
    'Τύπος παραστατικού
    If DisplayMessage(1, 4, 1, "", txtInvoiceCodeID.text) Then
        txtCodeShortDescription.SetFocus
        Exit Function
    End If
    
    'Νο παραστατικού
    If DisplayMessage(1, 4, 1, "", txtInvoiceNo.text) Then
        txtInvoiceNo.SetFocus
        Exit Function
    End If
    
    'Νο παραστατικού = ακέραιος
    If Not CheckForInteger(txtInvoiceNo.text) Then
        If DisplayMessage(2, 4, 1, "", txtInvoiceNo.text) Then
            txtInvoiceNo.SetFocus
            Exit Function
        End If
    End If
    
    'Είδη
    With grdItemsTransactions
        blnLineIsCorrect = False
        For lngRow = 1 To .RowCount
            blnSomethingIsGiven = False
            For lngCol = 1 To .colCount
                If .CellText(lngRow, lngCol) <> "" Then
                    blnSomethingIsGiven = True
                End If
            Next lngCol
            If blnSomethingIsGiven Then
                If .CellValue(lngRow, "ItemID") = "" Or .CellValue(lngRow, "Qty") = "" Then
                    If MyMsgBox(4, lblTitle.Caption, strMessages(11) & lngRow & " δεν είναι σωστή.", 1) Then
                    End If
                    .SetCurCell lngRow, "CategoryShortDescription"
                    .SetFocus
                    Exit Function
                Else
                    blnLineIsCorrect = True
                End If
            End If
        Next lngRow
        If Not blnLineIsCorrect Then
            If DisplayMessage(37, 4, 1, "", "") Then
            End If
            txtInvoiceRemarks.SetFocus
            Exit Function
        End If
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

