VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form UtilsPrinters 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   12525
   ClientLeft      =   15
   ClientTop       =   105
   ClientWidth     =   17400
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12525
   ScaleWidth      =   17400
   ShowInTaskbar   =   0   'False
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   5
      Left            =   10725
      TabIndex        =   65
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
      PicNormal       =   "UtilsPrinters.frx":0000
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newText txtOrientationDescription 
      Height          =   465
      Left            =   8100
      TabIndex        =   13
      Top             =   5325
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   820
      ForeColor       =   0
      Text            =   "Οριζόντιος"
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
      Index           =   4
      Left            =   10725
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   4800
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
      PicNormal       =   "UtilsPrinters.frx":059A
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newText txtPaperSizeDescription 
      Height          =   465
      Left            =   8100
      TabIndex        =   12
      Top             =   4800
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   820
      ForeColor       =   0
      Text            =   "A4 210mm x 297mm"
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
   Begin UserControls.newInteger mskPrinterReportHeight 
      Height          =   465
      Left            =   8100
      TabIndex        =   14
      Top             =   5850
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   2
      Text            =   "99"
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
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   75
      TabIndex        =   53
      Top             =   8775
      Width           =   7515
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   54
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
         TabIndex        =   55
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
         Index           =   4
         Left            =   5925
         TabIndex        =   56
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
         TabIndex        =   57
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
         TabIndex        =   58
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
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   3
      Left            =   8775
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   4275
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
      PicNormal       =   "UtilsPrinters.frx":0B34
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   2
      Left            =   4200
      TabIndex        =   49
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
      PicNormal       =   "UtilsPrinters.frx":10CE
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   1
      Left            =   4200
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   4275
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
      PicNormal       =   "UtilsPrinters.frx":1668
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin VB.Frame frmInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      DragIcon        =   "UtilsPrinters.frx":1C02
      DragMode        =   1  'Automatic
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   12375
      TabIndex        =   36
      Top             =   5025
      Width           =   4515
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   67
         TabStop         =   0   'False
         Text            =   "Orientations.OrientationID"
         Top             =   2700
         Width           =   3540
      End
      Begin VB.TextBox txtOrientationID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   66
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   2700
         Width           =   780
      End
      Begin VB.TextBox txtPaperSizeCodeNumber 
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
         TabIndex        =   63
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   2325
         Width           =   780
      End
      Begin VB.TextBox Text7 
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
         TabIndex        =   62
         TabStop         =   0   'False
         Text            =   "PaperSizes.PaperSizeCodeNumber"
         Top             =   2325
         Width           =   3540
      End
      Begin VB.TextBox txtCurrentGridRow 
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
         TabIndex        =   52
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   3075
         Width           =   780
      End
      Begin VB.TextBox Text6 
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
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "Grid.CurrentLine"
         Top             =   3075
         Width           =   3540
      End
      Begin VB.TextBox txtPrinterEAFDSSID 
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
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "4"
         Top             =   1200
         Width           =   780
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
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "Printers.PrinterEafdssID"
         Top             =   1200
         Width           =   3540
      End
      Begin VB.TextBox txtPrinterPrintsReportsID 
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
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "5"
         Top             =   1575
         Width           =   780
      End
      Begin VB.TextBox txtPrinterPrintsInvoicesID 
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
         TabIndex        =   43
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
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "Printers.PrinterPrintsReportsID"
         Top             =   1575
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
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "Printers.PrinterPrintsInvoicesID"
         Top             =   825
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
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "Printers.PrinterID"
         Top             =   75
         Width           =   3540
      End
      Begin VB.TextBox txtPrinterID 
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
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   75
         Width           =   780
      End
      Begin VB.TextBox txtPrinterTypeID 
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
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "Printers.PrinterTypeID"
         Top             =   450
         Width           =   3540
      End
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   75
         Top             =   3600
         _ExtentX        =   953
         _ExtentY        =   953
         IconSizeX       =   26
         IconSizeY       =   32
         Size            =   14064
         Images          =   "UtilsPrinters.frx":24CC
         Version         =   131072
         KeyCount        =   4
         Keys            =   "ÿÿÿ"
      End
   End
   Begin UserControls.newInteger mskPrinterInvoiceHeight 
      Height          =   465
      Left            =   3525
      TabIndex        =   8
      Top             =   6375
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   2
      Text            =   "99"
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
   Begin iGrid300_10Tec.iGrid grdPrinters 
      Height          =   3090
      Left            =   12000
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1050
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   5450
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
   Begin UserControls.newText txtPrinterName 
      Height          =   465
      Left            =   2100
      TabIndex        =   0
      Top             =   1125
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   40
      Text            =   "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ"
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
   Begin UserControls.newText txtPrinterEAFDSSString 
      Height          =   465
      Left            =   3525
      TabIndex        =   7
      Top             =   5850
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   0
      MaxLength       =   10
      Text            =   "ΑΑΑΑΑΑΑΑΑΑ"
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
   Begin UserControls.newInteger mskPrinterInvoiceDetailLines 
      Height          =   465
      Left            =   3525
      TabIndex        =   9
      Top             =   6900
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   2
      Text            =   "99"
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
   Begin UserControls.newInteger mskPrinterInvoiceTopMargin 
      Height          =   465
      Left            =   3525
      TabIndex        =   10
      Top             =   7425
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   2
      Text            =   "99"
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
   Begin UserControls.newInteger mskPrinterReportDetailLines 
      Height          =   465
      Left            =   8100
      TabIndex        =   15
      Top             =   6375
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   2
      Text            =   "99"
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
   Begin UserControls.newInteger mskPrinterReportTopMargin 
      Height          =   465
      Left            =   8100
      TabIndex        =   16
      Top             =   6900
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   2
      Text            =   "99"
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
   Begin UserControls.newInteger mskPrinterReportLeftMargin 
      Height          =   465
      Left            =   8100
      TabIndex        =   17
      Top             =   7425
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   2
      Text            =   "99"
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
   Begin UserControls.newInteger mskPrinterFontSize 
      Height          =   465
      Left            =   2100
      TabIndex        =   4
      Top             =   3225
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   2
      Text            =   "99"
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
   Begin UserControls.newText txtPrinterFontName 
      Height          =   465
      Left            =   2100
      TabIndex        =   3
      Top             =   2700
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   40
      Text            =   "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ"
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
   Begin iGrid300_10Tec.iGrid grdAvailablePrinters 
      Height          =   3540
      Left            =   12000
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4200
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   6244
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
   Begin UserControls.newText txtPrinterTypeDescription 
      Height          =   465
      Left            =   2100
      TabIndex        =   2
      Top             =   2175
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   40
      Text            =   "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ"
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
   Begin UserControls.newText txtPrintsReportsDescription 
      Height          =   465
      Left            =   8100
      TabIndex        =   11
      Top             =   4275
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   0
      Text            =   "ΝΑΙ"
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
   Begin UserControls.newText txtPrintsInvoicesDescription 
      Height          =   465
      Left            =   3525
      TabIndex        =   5
      Top             =   4275
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   0
      Text            =   "ΝΑΙ"
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
   Begin UserControls.newText txtEafdssDescription 
      Height          =   465
      Left            =   3525
      TabIndex        =   6
      Top             =   5325
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   0
      Text            =   "ΝΑΙ"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Αναφορές "
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   4515
      Index           =   1
      Left            =   5700
      TabIndex        =   26
      Tag             =   "SameColorAsBackground"
      Top             =   3750
      Width           =   5865
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   390
         Index           =   16
         Left            =   2475
         Top             =   4125
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Προσανατολισμός"
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
         Index           =   16
         Left            =   450
         TabIndex        =   64
         Top             =   1650
         Width           =   1515
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Μέγεθος χαρτιού"
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
         TabIndex        =   60
         Top             =   1125
         Width           =   1515
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Υψος σε γραμμές"
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
         TabIndex        =   59
         Top             =   2175
         Width           =   1515
         WordWrap        =   -1  'True
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   4
         Left            =   3450
         Top             =   675
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   9
         Left            =   5400
         Top             =   1050
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   8
         Left            =   1950
         Top             =   900
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   7
         Left            =   2475
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   3
         Left            =   0
         Top             =   450
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Εκτυπώνει αναφορές"
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
         TabIndex        =   35
         Top             =   600
         Width           =   1515
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Αναλυτικές γραμμές"
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
         TabIndex        =   34
         Top             =   2700
         Width           =   1515
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Επάνω περιθώριο"
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
         TabIndex        =   33
         Top             =   3225
         Width           =   1515
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Αριστερό περιθώριο"
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
         TabIndex        =   32
         Top             =   3750
         Width           =   1515
         WordWrap        =   -1  'True
      End
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   0
      Left            =   7125
      TabIndex        =   47
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
      PicNormal       =   "UtilsPrinters.frx":5BDC
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newText txtFriendlyName 
      Height          =   465
      Left            =   2100
      TabIndex        =   1
      Top             =   1650
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   40
      Text            =   "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Παραστατικά "
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   4515
      Index           =   0
      Left            =   450
      TabIndex        =   25
      Tag             =   "SameColorAsBackground"
      Top             =   3750
      Width           =   5190
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Με σήμανση"
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
         Index           =   18
         Left            =   450
         TabIndex        =   69
         Top             =   1650
         Width           =   2190
         WordWrap        =   -1  'True
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   2
         Left            =   4725
         Top             =   2100
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   1
         Left            =   2625
         Top             =   1050
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   0
         Left            =   3000
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   390
         Index           =   6
         Left            =   2775
         Top             =   3600
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   5
         Left            =   0
         Top             =   450
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Εκτυπώνει παραστατικά"
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
         TabIndex        =   31
         Top             =   600
         Width           =   1740
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Υψος σε γραμμές"
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
         TabIndex        =   30
         Top             =   2700
         Width           =   1740
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Αναλυτικές γραμμές"
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
         Top             =   3225
         Width           =   1740
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Συμβολοσειρά σήμανσης"
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
         TabIndex        =   28
         Top             =   2175
         Width           =   1740
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Επάνω περιθώριο"
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
         TabIndex        =   27
         Top             =   3750
         Width           =   1740
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Φιλική ονομασία"
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
      Index           =   17
      Left            =   450
      TabIndex        =   68
      Top             =   1725
      Width           =   1215
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   15
      Left            =   11550
      Top             =   4275
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   14
      Left            =   1650
      Top             =   1800
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
      Left            =   3900
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
      Top             =   3975
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   4200
      Top             =   9450
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   16650
      Top             =   2850
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   6900
      Top             =   8250
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   11
      Left            =   11700
      Top             =   -75
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   10
      Left            =   8550
      Top             =   3750
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Τύπος"
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
      TabIndex        =   24
      Top             =   2250
      Width           =   1065
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Εκτυπωτές"
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
      TabIndex        =   22
      Top             =   75
      Width           =   2475
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Γραμματοσειρά"
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
      TabIndex        =   20
      Top             =   2775
      Width           =   1065
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Μέγεθος"
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
      TabIndex        =   19
      Top             =   3300
      Width           =   1065
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Όνομα"
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
      TabIndex        =   18
      Top             =   1200
      Width           =   1065
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
   Begin VB.Menu mnuHdrPopUp 
      Caption         =   "mnuHdrPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuΑποθήκευσηΠλάτουςΣτηλών 
         Caption         =   "Αποθήκευση πλάτους στηλών"
      End
   End
End
Attribute VB_Name = "UtilsPrinters"
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
            ClearFields txtPrinterID, txtPrinterName, txtFriendlyName, txtPrinterTypeID, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrinterPrintsInvoicesID, txtPrintsInvoicesDescription, txtPrinterEAFDSSID, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrinterPrintsReportsID, txtPrintsReportsDescription, txtPaperSizeCodeNumber, txtPaperSizeDescription, txtOrientationID, txtOrientationDescription, mskPrinterReportHeight, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin
            DisableFields txtPrinterName, txtFriendlyName, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrintsInvoicesDescription, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrintsReportsDescription, mskPrinterReportHeight, txtPaperSizeDescription, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
            grdPrinters.SetFocus
            UpdateButtons Me, 4, 1, 0, 0, 0, 1
        End If
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Function DeleteRecord()
    
    If MainDeleteRecord("PrintersDB", "Printers", strAppTitle, "ID", txtPrinterID.text, "True") Then
        PopulateGrid
        HighlightNextRow grdPrinters, Val(txtCurrentGridRow.text), 2, True
        ClearFields txtPrinterID, txtPrinterName, txtFriendlyName, txtPrinterTypeID, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrinterPrintsInvoicesID, txtPrintsInvoicesDescription, txtPrinterEAFDSSID, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrinterPrintsReportsID, txtPrintsReportsDescription, txtPaperSizeCodeNumber, txtPaperSizeDescription, txtOrientationID, txtOrientationDescription, mskPrinterReportHeight, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin
        DisableFields txtPrinterName, txtFriendlyName, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrintsInvoicesDescription, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrintsReportsDescription, txtPaperSizeDescription, txtOrientationDescription, mskPrinterReportHeight, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
        UpdateButtons Me, 4, 1, 0, 0, 0, 1
    End If

End Function

Private Function FindPrinters()

    Dim prt As Printer
    Dim strSavedLayout As String
    
    With grdAvailablePrinters
        With .AddCol(sKey:="ID", sHeader:="ID", lWidth:=254, eHdrTextFlags:=igTextCenter)
            .eTextFlags = igTextCenter
        End With
        With .AddCol(sKey:="PrinterName", sHeader:="Όνομα", lWidth:=254, eHdrTextFlags:=igTextCenter)
            .eTextFlags = igTextLeft
        End With
    End With
    
    For Each prt In Printers
        grdAvailablePrinters.AddRow
        grdAvailablePrinters.CellValue(grdAvailablePrinters.RowCount, "PrinterName") = prt.DeviceName
    Next
    
    strSavedLayout = GetSetting(strAppTitle, "Layout Strings", "grdUtilsPrinters"): grdAvailablePrinters.LayoutCol = strSavedLayout

End Function

Private Function NewRecord()
    
    blnStatus = True
    ClearFields txtPrinterID, txtPrinterName, txtFriendlyName, txtPrinterTypeID, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrinterPrintsInvoicesID, txtPrintsInvoicesDescription, txtPrinterEAFDSSID, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrinterPrintsReportsID, txtPrintsReportsDescription, txtPaperSizeCodeNumber, txtPaperSizeDescription, txtOrientationID, txtOrientationDescription, mskPrinterReportHeight, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin
    EnableFields txtPrinterName, txtFriendlyName, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrintsInvoicesDescription, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrintsReportsDescription, txtPaperSizeDescription, txtOrientationDescription, mskPrinterReportHeight, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
    InitializeFields mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, mskPrinterFontSize, mskPrinterReportHeight, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin
    UpdateButtons Me, 4, 0, 1, 0, 1, 0
    txtPrinterName.SetFocus

End Function

Private Function PopulateGrid()

    If FillGridFromDB("PrintersDB", grdPrinters, "Printers", "", "", "", 2, 0, 2) Then
        grdPrinters.SetFocus
        grdPrinters.SetCurCell 1, 1
    End If

End Function

Private Function SaveRecord()
    
    Dim blnNotError
    
    If Not ValidateFields Then Exit Function
    
    blnNotError = MainSaveRecord("PrintersDB", "Printers", blnStatus, strAppTitle, "ID", txtPrinterID.text, txtPrinterName.text, txtFriendlyName.text, txtPrinterTypeID.text, txtPrinterFontName.text, mskPrinterFontSize.text, txtPrinterPrintsInvoicesID.text, txtPrinterEAFDSSID.text, txtPrinterEAFDSSString.text, mskPrinterInvoiceHeight.text, mskPrinterInvoiceDetailLines.text, mskPrinterInvoiceTopMargin.text, txtPrinterPrintsReportsID.text, IIf(txtPaperSizeCodeNumber.text <> "", txtPaperSizeCodeNumber.text, 0), IIf(txtOrientationID.text <> "", txtOrientationID.text, 0), mskPrinterReportHeight.text, mskPrinterReportDetailLines.text, mskPrinterReportTopMargin.text, mskPrinterReportLeftMargin.text, 1, strCurrentUser)
    
    If IsNumeric(blnNotError) And blnNotError Then
        txtPrinterID.text = blnNotError
        PopulateGrid
        HighlightRow grdPrinters, 1, txtPrinterID.text, True
        ClearFields txtPrinterID, txtPrinterName, txtFriendlyName, txtPrinterTypeID, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrinterPrintsInvoicesID, txtPrintsInvoicesDescription, txtPrinterEAFDSSID, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrinterPrintsReportsID, txtPrintsReportsDescription, txtPaperSizeCodeNumber, txtPaperSizeDescription, txtOrientationID, txtOrientationDescription, mskPrinterReportHeight, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin
        DisableFields txtPrinterName, txtFriendlyName, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrintsInvoicesDescription, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrintsReportsDescription, txtPaperSizeDescription, txtOrientationDescription, mskPrinterReportDetailLines, mskPrinterReportHeight, mskPrinterReportTopMargin, mskPrinterReportLeftMargin, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
        UpdateButtons Me, 4, 1, 0, 0, 0, 1
    End If

End Function

Private Function SeekRecord()
    
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    If grdPrinters.RowCount = 0 Then Exit Function
    
    ClearFields txtPrinterID, txtPrinterName, txtFriendlyName, txtPrinterTypeID, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrinterPrintsInvoicesID, txtPrintsInvoicesDescription, txtPrinterEAFDSSID, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrinterPrintsReportsID, txtPrintsReportsDescription, txtPaperSizeCodeNumber, txtPaperSizeDescription, txtOrientationID, txtOrientationDescription, mskPrinterReportHeight, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin
    DisableFields txtPrinterName, txtFriendlyName, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrintsInvoicesDescription, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrintsReportsDescription, txtPaperSizeDescription, txtOrientationDescription, mskPrinterReportHeight, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
    
    If MainSeekRecord("PrintersDB", "Printers", "ID", grdPrinters.CellValue(grdPrinters.CurRow, 1), True, txtPrinterID, txtPrinterName, txtFriendlyName, txtPrinterTypeID, txtPrinterFontName, mskPrinterFontSize, txtPrinterPrintsInvoicesID, txtPrinterEAFDSSID, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrinterPrintsReportsID, txtPaperSizeCodeNumber, txtOrientationID, mskPrinterReportHeight, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin) Then
        'Τύπος εκτυπωτή
        Set tmpRecordset = CheckForMatch("PrintersDB", txtPrinterTypeID.text, "PrinterTypes", "PrinterTypeID", "Numeric", 0, 1)
        txtPrinterTypeID.text = tmpRecordset.Fields(0)
        txtPrinterTypeDescription.text = tmpRecordset.Fields(1)
        'Εκτυπώνει παραστατικά;
        Set tmpRecordset = CheckForMatch("CommonDB", txtPrinterPrintsInvoicesID.text, "YesOrNo", "YesNoID", "Numeric", 0, 1)
        txtPrinterPrintsInvoicesID.text = tmpRecordset.Fields(0)
        txtPrintsInvoicesDescription.text = tmpRecordset.Fields(1)
        'Με σήμανση;
        Set tmpRecordset = CheckForMatch("CommonDB", txtPrinterEAFDSSID.text, "YesOrNo", "YesNoID", "String", 0, 1)
        txtPrinterEAFDSSID.text = tmpRecordset.Fields(0)
        txtEafdssDescription.text = tmpRecordset.Fields(1)
        'Εκτυπώνει αναφορές;
        Set tmpRecordset = CheckForMatch("CommonDB", txtPrinterPrintsReportsID.text, "YesOrNo", "YesNoID", "String", 0, 1)
        txtPrinterPrintsReportsID.text = tmpRecordset.Fields(0)
        txtPrintsReportsDescription.text = tmpRecordset.Fields(1)
        'Μέγεθος χαρτιού αναφορών
        Set tmpRecordset = CheckForMatch("PrintersDB", txtPaperSizeCodeNumber.text, "PaperSizes", "PaperSizeCodeNumber", "Numeric", 0, 1)
        If tmpRecordset.RecordCount > 0 Then
            txtPaperSizeDescription.text = tmpRecordset.Fields(1)
            txtPaperSizeCodeNumber.text = tmpRecordset.Fields(2)
        End If
        'Προσανατολισμός
        Set tmpRecordset = CheckForMatch("PrintersDB", txtOrientationID.text, "Orientations", "OrientationID", "Numeric", 0, 1)
        If tmpRecordset.RecordCount > 0 Then
            txtOrientationID.text = tmpRecordset.Fields(0)
            txtOrientationDescription.text = tmpRecordset.Fields(1)
        End If
        '
        blnStatus = False
        EnableFields txtPrinterName, txtFriendlyName, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrintsInvoicesDescription, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrintsReportsDescription, txtPaperSizeDescription, txtOrientationDescription, mskPrinterReportHeight, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
        UpdateButtons Me, 4, 0, 1, 1, 1, 0
        txtPrinterName.SetFocus
        txtCurrentGridRow.text = grdPrinters.CurRow
    End If
    
End Function

Private Function ValidateFields()

    ValidateFields = False
    
    'Περιγραφή
    If DisplayMessage(1, 4, 1, "", txtPrinterName.text) Then txtPrinterName.SetFocus: Exit Function
    
    'Τύπος
    If DisplayMessage(1, 4, 1, "", txtPrinterTypeID.text) Then txtPrinterTypeDescription.SetFocus: Exit Function
    
    'Εκτυπώνει παραστατικά
    If DisplayMessage(1, 4, 1, "", txtPrinterPrintsInvoicesID.text) Then txtPrintsInvoicesDescription.SetFocus: Exit Function
    
    'Εκτυπώνει παραστατικά με σήμανση
    If DisplayMessage(1, 4, 1, "", txtPrinterEAFDSSID.text) Then txtEafdssDescription.SetFocus: Exit Function
    
    'Υψος παραστατικού
    If DisplayMessage(1, 4, 1, "", mskPrinterInvoiceHeight.text) Then mskPrinterInvoiceHeight.SetFocus: Exit Function
    
    'Αναλυτικές γραμμές παραστατικού
    If DisplayMessage(1, 4, 1, "", mskPrinterInvoiceDetailLines.text) Then mskPrinterInvoiceDetailLines.SetFocus: Exit Function
    
    'Επάνω περιθώριο παραστατικού
    If DisplayMessage(1, 4, 1, "", mskPrinterInvoiceTopMargin.text) Then mskPrinterInvoiceTopMargin.SetFocus: Exit Function
    
    'Εκτυπώνει αναφορές
    If DisplayMessage(1, 4, 1, "", txtPrinterPrintsReportsID.text) Then txtPrintsReportsDescription.SetFocus: Exit Function
    
    'Μέγεθος χαρτιού αναφορών
    If txtPrinterPrintsReportsID.text = "1" Then If DisplayMessage(1, 4, 1, "", txtPaperSizeCodeNumber.text) Then txtPaperSizeDescription.SetFocus: Exit Function
        
    'Υψος αναφορών
    If DisplayMessage(1, 4, 1, "", mskPrinterReportHeight.text) Then mskPrinterReportHeight.SetFocus: Exit Function
    
    'Αναλυτικές γραμμές αναφορών
    If DisplayMessage(1, 4, 1, "", mskPrinterReportDetailLines.text) Then mskPrinterReportDetailLines.SetFocus: Exit Function
    
    'Επάνω περιθώριο αναφορών
    If DisplayMessage(1, 4, 1, "", mskPrinterReportTopMargin.text) Then mskPrinterReportTopMargin.SetFocus: Exit Function
    
    'Αριστερό περιθώριο αναφορών
    If DisplayMessage(1, 4, 1, "", mskPrinterReportLeftMargin.text) Then mskPrinterReportLeftMargin.SetFocus: Exit Function
    
    ValidateFields = True

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
            AbortProcedure False
        Case 4
            AbortProcedure True
    End Select

End Sub

Private Sub cmdIndex_Click(Index As Integer)
    
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case Index
        Case 0
            'Τύπος εκτυπωτή
            If txtPrinterTypeDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("PrintersDB", txtPrinterTypeDescription.text, "PrinterTypes", "PrinterTypeDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtPrinterTypeID.text = tmpTableData.strCode
            txtPrinterTypeDescription.text = tmpTableData.strOneField
        Case 1
            'Εκτυπώνει παραστατικά;
            If txtPrintsInvoicesDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtPrintsInvoicesDescription.text, "YesOrNo", "YesNoDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtPrinterPrintsInvoicesID.text = tmpTableData.strCode
            txtPrintsInvoicesDescription.text = tmpTableData.strOneField
        Case 2
            'Προτυπωμένο έντυπο
        Case 3
            'Με σήμανση;
            If txtEafdssDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtEafdssDescription.text, "YesOrNo", "YesNoDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtPrinterEAFDSSID.text = tmpTableData.strCode
            txtEafdssDescription.text = tmpTableData.strOneField
        Case 4
            'Εκτυπώνει αναφορές;
            If txtPrintsReportsDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtPrintsReportsDescription.text, "YesOrNo", "YesNoDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtPrinterPrintsReportsID.text = tmpTableData.strCode
            txtPrintsReportsDescription.text = tmpTableData.strOneField
        Case 5
            'Μέγεθος χαρτιού αναφορών
            If txtPaperSizeDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("PrintersDB", txtPaperSizeDescription.text, "PaperSizes", "PaperSizeDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 3, 0, 1, 2, "ID", "Περιγραφή", "Κωδικός μεγέθους", 0, 40, 10, 1, 0, 1)
            txtPaperSizeDescription.text = tmpTableData.strOneField
            txtPaperSizeCodeNumber.text = tmpTableData.strTwoField
        Case 6
            'Προσανατολισμός
            If txtOrientationDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("PrintersDB", txtOrientationDescription.text, "Orientations", "OrientationDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtOrientationID.text = tmpTableData.strCode
            txtOrientationDescription.text = tmpTableData.strOneField
    End Select

End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        FindPrinters
        Me.Tag = "False"
        AddColumnsToGrid grdPrinters, 25, GetSetting(strAppTitle, "Layout Strings", "grdUtilsPrinters"), "04NCNID,40NLNFriendlyName", "ID,Ονομασία"
        Me.Refresh
        PopulateGrid
    End If
    
    'AddDummyLines grdUtilsPrinters, "99999", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"

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
        Case vbKeyEscape
            If cmdButton(3).Enabled Then cmdButton_Click 3: Exit Function
            If cmdButton(4).Enabled Then cmdButton_Click 4
        Case vbKeyF12
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)

End Sub

Private Sub Form_Load()
    
    SetUpGrid lstIconList, grdPrinters: SetUpGrid lstIconList, grdAvailablePrinters
    PositionControls Me, False: ColorizeControls Me
    ClearFields txtPrinterID, txtFriendlyName, txtPrinterName, txtPrinterTypeID, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrinterPrintsInvoicesID, txtPrintsInvoicesDescription, txtPrinterEAFDSSID, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrinterPrintsReportsID, txtPrintsReportsDescription, txtPaperSizeCodeNumber, txtPaperSizeDescription, txtOrientationID, txtOrientationDescription, mskPrinterReportHeight, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin
    DisableFields txtPrinterName, txtFriendlyName, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrintsInvoicesDescription, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrintsReportsDescription, txtPaperSizeDescription, txtOrientationDescription, mskPrinterReportHeight, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
    UpdateButtons Me, 4, 1, 0, 0, 0, 1
    
End Sub

Private Sub grdPrinters_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    SeekRecord

End Sub


Private Sub grdPrinters_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp
    
End Sub


Private Sub grdPrinters_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SeekRecord
    
End Sub

Private Sub grdAvailablePrinters_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If txtPrinterName.Enabled Then
        txtPrinterName.text = grdAvailablePrinters.CellValue(lRow, "PrinterName")
        txtPrinterName.SetFocus
    End If

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()
    
    SaveSetting strAppTitle, "Layout Strings", "grdUtilsPrinters", grdPrinters.LayoutCol

End Sub

Private Sub txtEafdssDescription_Change()

    If txtEafdssDescription.text = "" Then ClearFields txtPrinterEAFDSSID

End Sub

Private Sub txtEafdssDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 3

End Sub

Private Sub txtEafdssDescription_Validate(Cancel As Boolean)

    If txtPrinterEAFDSSID.text = "" And txtEafdssDescription.text <> "" Then cmdIndex_Click 3: If txtPrinterEAFDSSID.text = "" Then Cancel = True

End Sub

Private Sub txtOrientationDescription_Change()

    If txtOrientationDescription.text = "" Then ClearFields txtOrientationID

End Sub

Private Sub txtOrientationDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 6

End Sub

Private Sub txtOrientationDescription_Validate(Cancel As Boolean)

    If txtOrientationID.text = "" And txtOrientationDescription.text <> "" Then cmdIndex_Click 6: If txtOrientationID.text = "" Then Cancel = True

End Sub

Private Sub txtPaperSizeDescription_Change()

    If txtPaperSizeDescription.text = "" Then ClearFields txtPaperSizeCodeNumber

End Sub

Private Sub txtPaperSizeDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 5

End Sub


Private Sub txtPaperSizeDescription_Validate(Cancel As Boolean)

    If txtPaperSizeCodeNumber.text = "" And txtPaperSizeDescription.text <> "" Then cmdIndex_Click 5: If txtPaperSizeCodeNumber.text = "" Then Cancel = True

End Sub

Private Sub txtPrintsInvoicesDescription_Change()

    If txtPrintsInvoicesDescription.text = "" Then ClearFields txtPrinterPrintsInvoicesID

End Sub

Private Sub txtPrintsInvoicesDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1

End Sub

Private Sub txtPrintsInvoicesDescription_Validate(Cancel As Boolean)

    If txtPrinterPrintsInvoicesID.text = "" And txtPrintsInvoicesDescription.text <> "" Then cmdIndex_Click 1: If txtPrinterPrintsInvoicesID.text = "" Then Cancel = True

End Sub

Private Sub txtPrintsReportsDescription_Change()

    If txtPrintsReportsDescription.text = "" Then ClearFields txtPrinterPrintsReportsID
    
End Sub

Private Sub txtPrintsReportsDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 4
    
End Sub

Private Sub txtPrintsReportsDescription_Validate(Cancel As Boolean)

    If txtPrinterPrintsReportsID.text = "" And txtPrintsReportsDescription.text <> "" Then cmdIndex_Click 4: If txtPrinterPrintsReportsID.text = "" Then Cancel = True
    
End Sub

Private Sub txtPrinterTypeDescription_Change()

    If txtPrinterTypeDescription.text = "" Then ClearFields txtPrinterTypeID

End Sub

Private Sub txtPrinterTypeDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub

Private Sub txtPrinterTypeDescription_Validate(Cancel As Boolean)

    If txtPrinterTypeID.text = "" And txtPrinterTypeDescription.text <> "" Then cmdIndex_Click 0: If txtPrinterTypeID.text = "" Then Cancel = True

End Sub

