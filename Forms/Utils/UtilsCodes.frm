VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form UtilsCodes 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   11910
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   17250
   ControlBox      =   0   'False
   ForeColor       =   &H00400000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11910
   ScaleWidth      =   17250
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   150
      TabIndex        =   52
      Top             =   8700
      Width           =   7515
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   53
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
         TabIndex        =   54
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
         TabIndex        =   55
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
         TabIndex        =   56
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
         TabIndex        =   57
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
      Height          =   4065
      Left            =   11775
      TabIndex        =   23
      Top             =   6000
      Width           =   4515
      Begin VB.TextBox Text9 
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
         TabIndex        =   75
         TabStop         =   0   'False
         Text            =   "Codes.CodeIsCreditID"
         Top             =   2700
         Width           =   3540
      End
      Begin VB.TextBox txtCodeIsCreditID 
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
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   2700
         Width           =   780
      End
      Begin VB.TextBox txtDateCheckID 
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
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   2325
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
         TabIndex        =   69
         TabStop         =   0   'False
         Text            =   "Codes.CodeDateCheck"
         Top             =   2325
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
         TabIndex        =   63
         TabStop         =   0   'False
         Text            =   "Codes.CodePrinterID"
         Top             =   1950
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
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   1950
         Width           =   780
      End
      Begin VB.TextBox txtIsPhysicalThingID 
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
         TabIndex        =   59
         TabStop         =   0   'False
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
         TabIndex        =   58
         TabStop         =   0   'False
         Text            =   "Codes.CodeIsPhysicalThingID"
         Top             =   1575
         Width           =   3540
      End
      Begin VB.TextBox Text7 
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
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "Codes.CodeTransformID"
         Top             =   1200
         Width           =   3540
      End
      Begin VB.TextBox txtTransformID 
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
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "Codes.CodeDetailsID"
         Top             =   450
         Width           =   3540
      End
      Begin VB.TextBox txtDetailsID 
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
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   450
         Width           =   780
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
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3075
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
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "RefersTo"
         Top             =   3075
         Width           =   3540
      End
      Begin VB.TextBox txtHandID 
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
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   825
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
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Codes.CodeHandID"
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
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "Codes.CodeID"
         Top             =   75
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
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   75
         Width           =   780
      End
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   75
         Top             =   3450
         _ExtentX        =   953
         _ExtentY        =   953
         IconSizeX       =   26
         IconSizeY       =   32
         Size            =   14064
         Images          =   "UtilsCodes.frx":0000
         Version         =   131072
         KeyCount        =   4
         Keys            =   "ÿÿÿ"
      End
   End
   Begin iGrid300_10Tec.iGrid grdUtilsCodes 
      Height          =   6615
      Left            =   9450
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1125
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   11668
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
   Begin UserControls.newText txtBatch 
      Height          =   465
      Left            =   3600
      TabIndex        =   3
      Top             =   2175
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
      MaxLength       =   4
      Text            =   "AAAA"
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
   Begin UserControls.newText txtShortDescription 
      Height          =   465
      Left            =   3600
      TabIndex        =   1
      Top             =   1125
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
      MaxLength       =   4
      Text            =   "ΑΑΑΑ"
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
   Begin UserControls.newText txtDescription 
      Height          =   465
      Left            =   3600
      TabIndex        =   2
      Top             =   1650
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   4194304
      MaxLength       =   40
      Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
   Begin UserControls.newText txtHandDescription 
      Height          =   465
      Left            =   7050
      TabIndex        =   7
      Top             =   3225
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
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
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   1
      Left            =   7725
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3225
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
      PicNormal       =   "UtilsCodes.frx":3710
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newText txtTransformDescription 
      Height          =   465
      Left            =   3600
      TabIndex        =   8
      Top             =   3750
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
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
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   3
      Left            =   4275
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3750
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
      PicNormal       =   "UtilsCodes.frx":3CAA
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newText txtDetailsDescription 
      Height          =   465
      Left            =   3600
      TabIndex        =   6
      Top             =   3225
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
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
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   0
      Left            =   4275
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3225
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
      PicNormal       =   "UtilsCodes.frx":4244
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newText txtCustomers 
      Height          =   465
      Left            =   3375
      TabIndex        =   15
      Top             =   6450
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
      MaxLength       =   1
      Text            =   "+"
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
   Begin UserControls.newText txtSuppliers 
      Height          =   465
      Left            =   4350
      TabIndex        =   16
      Top             =   6450
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
      MaxLength       =   1
      Text            =   "+"
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
   Begin UserControls.newInteger txtLastNo 
      Height          =   465
      Left            =   3600
      TabIndex        =   17
      Top             =   7725
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   820
      MaxLength       =   7
      Text            =   "999.999"
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
   Begin UserControls.newDate mskLastDate 
      Height          =   465
      Left            =   4500
      TabIndex        =   18
      Top             =   7725
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   820
      Text            =   "01/01/2017"
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
   Begin UserControls.newText txtInventoryQty 
      Height          =   465
      Left            =   1050
      TabIndex        =   13
      Top             =   6450
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
      MaxLength       =   1
      Text            =   "+"
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
   Begin UserControls.newText txtInventoryValue 
      Height          =   465
      Left            =   1950
      TabIndex        =   14
      Top             =   6450
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
      MaxLength       =   1
      Text            =   "+"
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
   Begin VB.Frame frmFrame 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Ενημερώνει "
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   2040
      Index           =   3
      Left            =   450
      TabIndex        =   40
      Tag             =   "SameColorAsBackground"
      Top             =   5325
      Width           =   5115
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   7
         Left            =   4650
         Top             =   750
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   11
         Left            =   0
         Top             =   750
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   12
         Left            =   2325
         Top             =   750
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "(+/-)"
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
         Index           =   22
         Left            =   450
         TabIndex        =   51
         Top             =   750
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "(+/-)"
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
         Index           =   21
         Left            =   2775
         TabIndex        =   50
         Top             =   750
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "(+/-)"
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
         Index           =   20
         Left            =   3750
         TabIndex        =   49
         Top             =   750
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "(+/-)"
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
         Index           =   19
         Left            =   1425
         TabIndex        =   48
         Top             =   750
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Πελάτες"
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
         Left            =   2775
         TabIndex        =   44
         Top             =   450
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Προμηθευτές"
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
         Left            =   3750
         TabIndex        =   43
         Top             =   450
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Αξίες"
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
         Index           =   15
         Left            =   1425
         TabIndex        =   42
         Top             =   450
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Ποσότητες"
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
         TabIndex        =   41
         Top             =   450
         Width           =   915
         WordWrap        =   -1  'True
      End
   End
   Begin UserControls.newText txtIsPhysicalThingDescription 
      Height          =   465
      Left            =   7050
      TabIndex        =   9
      Top             =   3750
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
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
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   4
      Left            =   7725
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   3750
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
      PicNormal       =   "UtilsCodes.frx":47DE
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newText txtPrinterDescription 
      Height          =   465
      Left            =   3600
      TabIndex        =   10
      Top             =   4275
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   4194304
      MaxLength       =   40
      Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
      Left            =   8625
      TabIndex        =   64
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
      PicNormal       =   "UtilsCodes.frx":4D78
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newInteger mskDetailLines 
      Height          =   465
      Left            =   7050
      TabIndex        =   5
      Top             =   2700
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
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
   Begin UserControls.newText txtDateCheckDescription 
      Height          =   465
      Left            =   3600
      TabIndex        =   4
      Top             =   2700
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
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
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   5
      Left            =   4275
      TabIndex        =   67
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
      PicNormal       =   "UtilsCodes.frx":5312
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newText txtCodeDTable 
      Height          =   465
      Left            =   3600
      TabIndex        =   11
      Top             =   4800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
      MaxLength       =   3
      Text            =   "ΑΑΑ"
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
   Begin UserControls.newText txtCodeIsCreditDescription 
      Height          =   465
      Left            =   7050
      TabIndex        =   12
      Top             =   4800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
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
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   6
      Left            =   7725
      TabIndex        =   72
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
      PicNormal       =   "UtilsCodes.frx":58AC
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Είναι ακυρωτικό"
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
      Left            =   5100
      TabIndex        =   73
      Top             =   4875
      Width           =   1515
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Κωδικός πίνακα Δέλτα"
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
      TabIndex        =   71
      Top             =   4875
      Width           =   2715
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Ελεγχος ημερομηνίας έκδοσης"
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
      TabIndex        =   68
      Top             =   2775
      Width           =   2715
      WordWrap        =   -1  'True
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   765
      Index           =   9
      Left            =   4200
      Top             =   2175
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   765
      Index           =   8
      Left            =   4650
      Top             =   3300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Πλήθος γραμμών"
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
      Left            =   5100
      TabIndex        =   66
      Top             =   2775
      Width           =   1515
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   6
      Left            =   6600
      Top             =   3225
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Εκτυπωτής"
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
      TabIndex        =   65
      Top             =   4350
      Width           =   2715
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Αφορά φυσικά είδη"
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
      Left            =   5100
      TabIndex        =   61
      Top             =   3825
      Width           =   1515
      WordWrap        =   -1  'True
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   5
      Left            =   3150
      Top             =   7650
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Τελευταίο παραστατικό που εκδόθηκε"
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
      TabIndex        =   47
      Top             =   7800
      Width           =   2715
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Index           =   10
      Left            =   4500
      TabIndex        =   46
      Top             =   7425
      Width           =   1515
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Νο"
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
      Left            =   3600
      TabIndex        =   45
      Top             =   7425
      Width           =   840
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Να ζητούνται τα 'Λοιπά στοιχεία'"
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
      TabIndex        =   32
      Top             =   3300
      Width           =   2715
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Μετασχηματίζεται σε ανώτερο επίπεδο"
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
      TabIndex        =   31
      Top             =   3825
      Width           =   2715
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   5700
      Top             =   8175
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   15150
      Top             =   4050
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   6525
      Top             =   9375
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   2
      Left            =   9000
      Top             =   3600
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
      Left            =   3150
      Top             =   2925
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
      Left            =   0
      Top             =   1650
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   3
      Left            =   3675
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Χειρόγραφο"
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
      Left            =   5100
      TabIndex        =   26
      Top             =   3300
      Width           =   1515
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Τύποι παραστατικών"
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
      Width           =   4815
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Σειρά"
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
      TabIndex        =   20
      Top             =   2250
      Width           =   2715
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Συντομογραφία"
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
      TabIndex        =   0
      Top             =   1200
      Width           =   2715
   End
   Begin VB.Label lblLabel 
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
      TabIndex        =   19
      Top             =   1725
      Width           =   2715
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
      Index           =   4
      Left            =   10950
      Top             =   -75
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
Attribute VB_Name = "UtilsCodes"
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
            ClearFields txtID, txtDetailsID, txtHandID, txtTransformID, txtIsPhysicalThingID, txtPrinterID, txtDateCheckID, txtCodeIsCreditID
            ClearFields txtShortDescription, txtDescription, txtBatch, txtDetailsDescription, txtHandDescription, txtTransformDescription, txtIsPhysicalThingDescription, txtInventoryQty, txtInventoryValue, txtCustomers, txtSuppliers, txtLastNo, mskLastDate, txtPrinterDescription, mskDetailLines, txtDateCheckDescription, txtCodeDTable, txtCodeIsCreditDescription
            DisableFields txtShortDescription, txtDescription, txtBatch, txtDetailsDescription, txtHandDescription, txtTransformDescription, txtIsPhysicalThingDescription, txtInventoryQty, txtInventoryValue, txtCustomers, txtSuppliers, txtLastNo, mskLastDate, txtPrinterDescription, mskDetailLines, txtDateCheckDescription, txtCodeDTable, txtCodeIsCreditDescription
            DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6)
            grdUtilsCodes.SetFocus
            UpdateButtons Me, 4, 1, 0, 0, 0, 1
        End If
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Function DeleteRecord()
    
    If MainDeleteRecord("CommonDB", "Codes", strAppTitle, "ID", Val(txtID.text), "True") Then
        PopulateGrid
        HighlightRow grdUtilsCodes, 1, "", True
        ClearFields txtID, txtDetailsID, txtHandID, txtTransformID, txtIsPhysicalThingID, txtPrinterID, txtDateCheckID, txtCodeIsCreditID
        ClearFields txtShortDescription, txtDescription, txtBatch, txtDetailsDescription, txtHandDescription, txtTransformDescription, txtIsPhysicalThingDescription, txtInventoryQty, txtInventoryValue, txtCustomers, txtSuppliers, txtLastNo, mskLastDate, txtPrinterDescription, mskDetailLines, txtDateCheckDescription, txtCodeDTable, txtCodeIsCreditDescription
        DisableFields txtShortDescription, txtDescription, txtBatch, txtDetailsDescription, txtHandDescription, txtTransformDescription, txtIsPhysicalThingDescription, txtInventoryQty, txtInventoryValue, txtCustomers, txtSuppliers, txtLastNo, mskLastDate, txtPrinterDescription, mskDetailLines, txtDateCheckDescription, txtCodeDTable, txtCodeIsCreditDescription
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6)
        UpdateButtons Me, 4, 1, 0, 0, 0, 1
    End If

End Function

Private Function PopulateGrid()
        
    If FillGridFromDB("CommonDB", grdUtilsCodes, "Codes", "", "", "CodeRefersTo = " & txtRefersTo.text, 4, 0, 1, 2, 8) Then
        grdUtilsCodes.SetFocus
        grdUtilsCodes.SetCurCell 1, 1
    End If

End Function

Private Function NewRecord()
    
    blnStatus = True
    
    ClearFields txtID, txtDetailsID, txtHandID, txtTransformID, txtIsPhysicalThingID, txtPrinterID, txtDateCheckID, txtCodeIsCreditID
    ClearFields txtShortDescription, txtDescription, txtBatch, txtDetailsDescription, txtHandDescription, txtTransformDescription, txtIsPhysicalThingDescription, txtInventoryQty, txtInventoryValue, txtCustomers, txtSuppliers, txtLastNo, mskLastDate, txtPrinterDescription, mskDetailLines, txtDateCheckDescription, txtCodeDTable, txtCodeIsCreditDescription
    EnableFields txtShortDescription, txtDescription, txtBatch, txtDetailsDescription, txtHandDescription, txtTransformDescription, txtIsPhysicalThingDescription, txtInventoryQty, txtInventoryValue, txtCustomers, txtSuppliers, txtLastNo, mskLastDate, txtPrinterDescription, mskDetailLines, txtDateCheckDescription, txtCodeDTable, txtCodeIsCreditDescription
    EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6)
    UpdateButtons Me, 4, 0, 1, 0, 1, 0
    InitializeFields txtLastNo, mskLastDate
    txtShortDescription.SetFocus

End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    If MainSaveRecord("CommonDB", "Codes", blnStatus, strAppTitle, "ID", _
        Val(txtID.text), _
        txtShortDescription.text, _
        txtDescription.text, _
        txtRefersTo.text, _
        txtInventoryQty.text, _
        txtInventoryValue.text, _
        txtCustomers.text, _
        txtSuppliers.text, _
        txtBatch.text, _
        txtDetailsID.text, _
        txtHandID.text, _
        IIf(txtPrinterID.text = "", "0", txtPrinterID.text), _
        txtCodeDTable.text, _
        txtCodeIsCreditID, _
        txtTransformID.text, _
        txtIsPhysicalThingID.text, _
        txtLastNo.text, _
        mskLastDate.text, _
        mskDetailLines.text, _
        txtDateCheckID.text, _
        txtRefersTo.text, _
        strCurrentUser) <> 0 Then
        PopulateGrid
        HighlightRow grdUtilsCodes, 1, txtID.text, True
        ClearFields txtID, txtDetailsID, txtHandID, txtTransformID, txtIsPhysicalThingID, txtPrinterID, txtDateCheckID, txtCodeIsCreditID
        ClearFields txtShortDescription, txtDescription, txtBatch, txtDetailsDescription, txtHandDescription, txtTransformDescription, txtIsPhysicalThingDescription, txtInventoryQty, txtInventoryValue, txtCustomers, txtSuppliers, txtLastNo, mskLastDate, txtPrinterDescription, mskDetailLines, txtDateCheckDescription, txtCodeDTable, txtCodeIsCreditDescription
        DisableFields txtShortDescription, txtDescription, txtBatch, txtDetailsDescription, txtHandDescription, txtTransformDescription, txtIsPhysicalThingDescription, txtInventoryQty, txtInventoryValue, txtCustomers, txtSuppliers, txtLastNo, mskLastDate, txtPrinterDescription, mskDetailLines, txtDateCheckDescription, txtCodeDTable, txtCodeIsCreditDescription
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6)
        UpdateButtons Me, 4, 1, 0, 0, 0, 1
    End If

End Function

Private Function SeekRecord()
    
    On Error GoTo ErrTrap
    
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Dim blnEnableDelete As Boolean
    
    If grdUtilsCodes.RowCount = 0 Then Exit Function
    
    ClearFields txtID, txtDetailsID, txtHandID, txtTransformID, txtIsPhysicalThingID, txtPrinterID, txtDateCheckID, txtCodeIsCreditID
    ClearFields txtShortDescription, txtDescription, txtBatch, txtDetailsDescription, txtHandDescription, txtTransformDescription, txtIsPhysicalThingDescription, txtInventoryQty, txtInventoryValue, txtCustomers, txtSuppliers, txtLastNo, mskLastDate, txtPrinterDescription, mskDetailLines, txtDateCheckDescription, txtCodeDTable, txtCodeIsCreditDescription
    DisableFields txtShortDescription, txtDescription, txtBatch, txtDetailsDescription, txtHandDescription, txtTransformDescription, txtIsPhysicalThingDescription, txtInventoryQty, txtInventoryValue, txtCustomers, txtSuppliers, txtLastNo, mskLastDate, txtPrinterDescription, mskDetailLines, txtDateCheckDescription, txtCodeDTable, txtCodeIsCreditDescription
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6)
    blnEnableDelete = SimpleSeek("Invoices", "InvoiceID", grdUtilsCodes.CellValue(grdUtilsCodes.CurRow, 1))
    
    If MainSeekRecord("CommonDB", "Codes", "ID", grdUtilsCodes.CellValue(grdUtilsCodes.CurRow, 1), True, txtID, _
        txtShortDescription, _
        txtDescription, _
        txtRefersTo, _
        txtInventoryQty, _
        txtInventoryValue, _
        txtCustomers, _
        txtSuppliers, _
        txtBatch, _
        txtDetailsID, _
        txtHandID, _
        txtPrinterID, _
        txtCodeDTable, _
        txtCodeIsCreditID, _
        txtTransformID, _
        txtIsPhysicalThingID, _
        txtLastNo, _
        mskLastDate, _
        mskDetailLines, _
        txtDateCheckID) Then
        'Να ζητούνται τα λοιπά στοιχεία
        Set tmpRecordset = CheckForMatch("CommonDB", txtDetailsID.text, "YesOrNo", "YesNoID", "Numeric", 0, 1)
        txtDetailsID.text = tmpRecordset.Fields(0)
        txtDetailsDescription.text = tmpRecordset.Fields(1)
        'Χειρόγραφο
        Set tmpRecordset = CheckForMatch("CommonDB", txtHandID.text, "YesOrNo", "YesNoID", "Numeric", 0, 1)
        txtHandID.text = tmpRecordset.Fields(0)
        txtHandDescription.text = tmpRecordset.Fields(1)
        'Εκτυπωτής
        Set tmpRecordset = CheckForMatch("PrintersDB", txtPrinterID.text, "Printers", "PrinterID", "Numeric", 0, 1)
        If tmpRecordset.RecordCount > 0 Then
            txtPrinterID.text = tmpRecordset.Fields(0)
            txtPrinterDescription.text = tmpRecordset.Fields(2)
        End If
        'Μετασχηματίζεται
        Set tmpRecordset = CheckForMatch("CommonDB", txtTransformID.text, "YesOrNo", "YesNoID", "Numeric", 0, 1)
        txtTransformID.text = tmpRecordset.Fields(0)
        txtTransformDescription.text = tmpRecordset.Fields(1)
        'Αφορά φυσικό είδος
        Set tmpRecordset = CheckForMatch("CommonDB", txtIsPhysicalThingID.text, "YesOrNo", "YesNoID", "Numeric", 0, 1)
        txtIsPhysicalThingID.text = tmpRecordset.Fields(0)
        txtIsPhysicalThingDescription.text = tmpRecordset.Fields(1)
        'Ελεγχος ημερομηνίας έκδοσης
        Set tmpRecordset = CheckForMatch("CommonDB", txtDateCheckID.text, "YesOrNo", "YesNoID", "Numeric", 0, 1)
        txtDateCheckID.text = tmpRecordset.Fields(0)
        txtDateCheckDescription.text = tmpRecordset.Fields(1)
        'Είναι ακυρωτικό
        Set tmpRecordset = CheckForMatch("CommonDB", txtCodeIsCreditID.text, "YesOrNo", "YesNoID", "Numeric", 0, 1)
        txtCodeIsCreditID.text = tmpRecordset.Fields(0)
        txtCodeIsCreditDescription.text = tmpRecordset.Fields(1)
        '
        blnStatus = False
        
        EnableFields txtShortDescription, txtDescription, txtBatch, txtDetailsDescription, txtHandDescription, txtTransformDescription, txtIsPhysicalThingDescription, txtInventoryQty, txtInventoryValue, txtCustomers, txtSuppliers, txtLastNo, mskLastDate, txtPrinterDescription, mskDetailLines, txtDateCheckDescription, txtCodeDTable, txtCodeIsCreditDescription
        EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6)
        
        UpdateButtons Me, 4, 0, 1, IIf(blnEnableDelete, 1, 0), 1, 0
        txtShortDescription.SetFocus
    End If
    
    Exit Function
    
ErrTrap:
    ClearFields txtID, txtDetailsID, txtHandID, txtTransformID, txtIsPhysicalThingID, txtPrinterID, txtDateCheckID, txtCodeIsCreditID
    ClearFields txtShortDescription, txtDescription, txtBatch, txtDetailsDescription, txtHandDescription, txtTransformDescription, txtIsPhysicalThingDescription, txtInventoryQty, txtInventoryValue, txtCustomers, txtSuppliers, txtLastNo, mskLastDate, txtPrinterDescription, mskDetailLines, txtDateCheckDescription, txtCodeDTable, txtCodeIsCreditDescription

    UpdateButtons Me, 4, 1, 0, 0, 0, 1
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function ValidateFields()

    ValidateFields = False
    
    'Συντομογραφία
    If DisplayMessage(1, 4, 1, "", txtShortDescription.text) Then txtShortDescription.SetFocus: Exit Function
    
    'Περιγραφή
    If DisplayMessage(1, 4, 1, "", txtDescription.text) Then txtDescription.SetFocus: Exit Function
    
    'Λοιπά στοιχεία
    If DisplayMessage(1, 4, 1, "", txtDetailsDescription.text) Then txtDetailsDescription.SetFocus: Exit Function
    
    'Χειρόγραφο
    If DisplayMessage(1, 4, 1, "", txtHandDescription.text) Then txtHandDescription.SetFocus: Exit Function
    
    'Μετασχηματίζεται
    If DisplayMessage(1, 4, 1, "", txtTransformDescription.text) Then txtTransformDescription.SetFocus: Exit Function
    
    'Είναι ακυρωτικό
    If DisplayMessage(1, 4, 1, "", txtCodeIsCreditDescription.text) Then txtCodeIsCreditDescription.SetFocus: Exit Function
    
    'Τελευταίο παραστατικό
    If DisplayMessage(1, 4, 1, "", txtLastNo.text) Then txtLastNo.SetFocus: Exit Function
    
    'Τελευταία ημερομηνία
    If Not IsDate(mskLastDate.text) Then
        If MyMsgBox(4, strAppTitle, strMessages(2), 1) Then
        End If
        mskLastDate.SetFocus
        Exit Function
    End If
    
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

    'Local μεταβλητές
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case Index
        Case 0
            'Να ζητούνται τα λοιπά στοιχεία
            If txtDetailsDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtDetailsDescription.text, "YesOrNo", "YesNoDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtDetailsID.text = tmpTableData.strCode
            txtDetailsDescription.text = tmpTableData.strOneField
        Case 1
            'Χειρόγραφο
            If txtHandDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtHandDescription.text, "YesOrNo", "YesNoDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtHandID.text = tmpTableData.strCode
            txtHandDescription.text = tmpTableData.strOneField
        Case 2
            'Εκτυπωτής
            If txtPrinterDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("PrintersDB", txtPrinterDescription.text, "Printers", "PrinterFriendlyName", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 2, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtPrinterID.text = tmpTableData.strCode
            txtPrinterDescription.text = tmpTableData.strOneField
        Case 3
            'Μετασχηματίζεται
            If txtTransformDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtTransformDescription.text, "YesOrNo", "YesNoDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtTransformID.text = tmpTableData.strCode
            txtTransformDescription.text = tmpTableData.strOneField
        Case 4
            'Αφορά φυσικό είδος
            If txtIsPhysicalThingDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtIsPhysicalThingDescription.text, "YesOrNo", "YesNoDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtIsPhysicalThingID.text = tmpTableData.strCode
            txtIsPhysicalThingDescription.text = tmpTableData.strOneField
        Case 5
            'Ελεγχος ημερομηνίας έκδοσης
            If txtDateCheckDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtDateCheckDescription.text, "YesOrNo", "YesNoDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtDateCheckID.text = tmpTableData.strCode
            txtDateCheckDescription.text = tmpTableData.strOneField
        Case 6
            'Είναι ακυρωτικό
            If txtCodeIsCreditDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtCodeIsCreditDescription.text, "YesOrNo", "YesNoDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtCodeIsCreditID.text = tmpTableData.strCode
            txtCodeIsCreditDescription.text = tmpTableData.strOneField
    End Select

End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdUtilsCodes, 25, GetSetting(strAppTitle, "Layout Strings", "grdUtilsCodes"), "04NCNID,04NCNShortDescription,40NLNDescription,04NCNBatch", "ID,Συντ.,Περιγραφή,Σειρά"
        Me.Refresh
        PopulateGrid
    End If

    'AddDummyLines grdUtilsCodes, 5, 5, 40, 4

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
    
    PositionControls Me, False: ColorizeControls Me
    SetUpGrid lstIconList, grdUtilsCodes
    ClearFields txtID, txtDetailsID, txtHandID, txtTransformID, txtIsPhysicalThingID, txtPrinterID, txtDateCheckID, txtCodeIsCreditID
    ClearFields txtShortDescription, txtDescription, txtBatch, txtDetailsDescription, txtHandDescription, txtTransformDescription, txtIsPhysicalThingDescription, txtInventoryQty, txtInventoryValue, txtCustomers, txtSuppliers, txtLastNo, mskLastDate, txtPrinterDescription, mskDetailLines, txtDateCheckDescription, txtCodeDTable, txtCodeIsCreditDescription
    DisableFields txtShortDescription, txtDescription, txtBatch, txtDetailsDescription, txtHandDescription, txtTransformDescription, txtIsPhysicalThingDescription, txtInventoryQty, txtInventoryValue, txtCustomers, txtSuppliers, txtLastNo, mskLastDate, txtPrinterDescription, mskDetailLines, txtDateCheckDescription, txtCodeDTable, txtCodeIsCreditDescription
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6)
    UpdateButtons Me, 4, 1, 0, 0, 0, 1
    
End Sub

Private Sub grdUtilsCodes_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    SeekRecord

End Sub

Private Sub grdUtilsCodes_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdUtilsCodes_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SeekRecord

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdUtilsCodes", grdUtilsCodes.LayoutCol

End Sub

Private Sub txtCodeIsCreditDescription_Change()

    If txtCodeIsCreditDescription.text = "" Then ClearFields txtCodeIsCreditID
    
End Sub

Private Sub txtCodeIsCreditDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 6
    
End Sub


Private Sub txtCodeIsCreditDescription_Validate(Cancel As Boolean)

    If txtCodeIsCreditID.text = "" And txtCodeIsCreditDescription.text <> "" Then cmdIndex_Click 6: If txtCodeIsCreditID.text = "" Then Cancel = True

End Sub

Private Sub txtCustomers_Change()

        If Not OnlyAcceptSpecificValues(txtCustomers.text, "+", "-") Then ClearFields txtCustomers

End Sub

Private Sub txtDateCheckDescription_Change()

    If txtDateCheckDescription.text = "" Then ClearFields txtDateCheckID

End Sub

Private Sub txtDateCheckDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 5
    
End Sub


Private Sub txtDateCheckDescription_Validate(Cancel As Boolean)

    If txtDateCheckID.text = "" And txtDateCheckDescription.text <> "" Then cmdIndex_Click 5: If txtDateCheckID.text = "" Then Cancel = True

End Sub

Private Sub txtDetailsDescription_Change()

    If txtDetailsDescription.text = "" Then ClearFields txtDetailsID

End Sub

Private Sub txtDetailsDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub


Private Sub txtDetailsDescription_Validate(Cancel As Boolean)

    If txtDetailsID.text = "" And txtDetailsDescription.text <> "" Then cmdIndex_Click 0: If txtDetailsID.text = "" Then Cancel = True

End Sub


Private Sub txtHandDescription_Change()

    If txtHandDescription.text = "" Then ClearFields txtHandID

End Sub

Private Sub txtHandDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1
    
End Sub

Private Sub txtHandDescription_Validate(Cancel As Boolean)

    If txtHandID.text = "" And txtHandDescription.text <> "" Then cmdIndex_Click 1: If txtHandID.text = "" Then Cancel = True

End Sub


Private Sub txtInventoryQty_Change()

    If Not OnlyAcceptSpecificValues(txtInventoryQty.text, "+", "-") Then ClearFields txtInventoryQty
    
End Sub

Private Sub txtInventoryValue_Change()

    If Not OnlyAcceptSpecificValues(txtInventoryValue.text, "+", "-") Then ClearFields txtInventoryValue
    
End Sub

Private Sub txtIsPhysicalThingDescription_Change()

    If txtIsPhysicalThingDescription.text = "" Then ClearFields txtIsPhysicalThingID

End Sub

Private Sub txtIsPhysicalThingDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 4

End Sub


Private Sub txtIsPhysicalThingDescription_Validate(Cancel As Boolean)

    If txtIsPhysicalThingID.text = "" And txtIsPhysicalThingDescription.text <> "" Then cmdIndex_Click 4: If txtIsPhysicalThingID.text = "" Then Cancel = True

End Sub

Private Sub txtPrinterDescription_Change()

    If txtPrinterDescription.text = "" Then ClearFields txtPrinterID

End Sub

Private Sub txtPrinterDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2

End Sub


Private Sub txtPrinterDescription_Validate(Cancel As Boolean)

    If txtPrinterID.text = "" And txtPrinterDescription.text <> "" Then cmdIndex_Click 2: If txtPrinterID.text = "" Then Cancel = True

End Sub

Private Sub txtSuppliers_Change()

        If Not OnlyAcceptSpecificValues(txtSuppliers.text, "+", "-") Then ClearFields txtSuppliers

End Sub

Private Sub txtTransformDescription_Change()

    If txtTransformDescription.text = "" Then ClearFields txtTransformID

End Sub


Private Sub txtTransformDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 3

End Sub


Private Sub txtTransformDescription_Validate(Cancel As Boolean)

    If txtTransformID.text = "" And txtTransformDescription.text <> "" Then cmdIndex_Click 3: If txtTransformID.text = "" Then Cancel = True

End Sub


