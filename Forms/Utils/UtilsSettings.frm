VERSION 5.00
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form UtilsSettings 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   13830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22545
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13830
   ScaleWidth      =   22545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmFrame 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   6165
      Index           =   0
      Left            =   11925
      TabIndex        =   32
      Top             =   6975
      Width           =   9165
      Begin VB.Frame Frame 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   " Επικεφαλίδες παραστατικών "
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3990
         Index           =   0
         Left            =   450
         TabIndex        =   33
         Top             =   450
         Width           =   8265
         Begin UserControls.newText txtLine01 
            Height          =   465
            Left            =   1650
            TabIndex        =   0
            Top             =   525
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin UserControls.newText txtLine02 
            Height          =   465
            Left            =   1650
            TabIndex        =   1
            Top             =   1050
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin UserControls.newText txtLine03 
            Height          =   465
            Left            =   1650
            TabIndex        =   2
            Top             =   1575
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin UserControls.newText txtLine04 
            Height          =   465
            Left            =   1650
            TabIndex        =   3
            Top             =   2100
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin UserControls.newText txtLine05 
            Height          =   465
            Left            =   1650
            TabIndex        =   4
            Top             =   2625
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin UserControls.newText txtLine06 
            Height          =   465
            Left            =   1650
            TabIndex        =   5
            Top             =   3150
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
            BackColor       =   &H000080FF&
            Caption         =   "1η Γραμμή"
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
            TabIndex        =   39
            Top             =   600
            Width           =   750
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "2η Γραμμή"
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
            TabIndex        =   38
            Top             =   1125
            Width           =   750
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "3η Γραμμή"
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
            TabIndex        =   37
            Top             =   1650
            Width           =   750
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "4η Γραμμή"
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
            Top             =   2175
            Width           =   750
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "5η Γραμμή"
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
            TabIndex        =   35
            Top             =   2700
            Width           =   750
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "6η Γραμμή"
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
            TabIndex        =   34
            Top             =   3225
            Width           =   750
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   9
            Left            =   0
            Top             =   1950
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   540
            Index           =   10
            Left            =   2550
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
            Index           =   11
            Left            =   2475
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
            Index           =   0
            Left            =   1200
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
            Index           =   13
            Left            =   7800
            Top             =   1650
            Visible         =   0   'False
            Width           =   465
         End
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   8
         Left            =   0
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
         Index           =   1
         Left            =   8700
         Top             =   1275
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin VB.Frame frmFrame 
      BackColor       =   &H00004080&
      BorderStyle     =   0  'None
      Height          =   6690
      Index           =   2
      Left            =   12975
      TabIndex        =   50
      Top             =   2925
      Width           =   9165
      Begin UserControls.newText txtPrintHourDescription 
         Height          =   465
         Left            =   4500
         TabIndex        =   10
         Top             =   450
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
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   1
         Left            =   5175
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   450
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
         PicNormal       =   "UtilsSettings.frx":0000
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newText txtPrintBalanceDescription 
         Height          =   465
         Left            =   4500
         TabIndex        =   11
         Top             =   975
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
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   2
         Left            =   5175
         TabIndex        =   56
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
         PicNormal       =   "UtilsSettings.frx":059A
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newText txtRoundSalesDescription 
         Height          =   465
         Left            =   4500
         TabIndex        =   12
         Top             =   1500
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
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   3
         Left            =   5175
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1500
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
         PicNormal       =   "UtilsSettings.frx":0B34
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newInteger mskRoundSalesCents 
         Height          =   465
         Left            =   4500
         TabIndex        =   13
         Top             =   2025
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
      Begin UserControls.newText txtTransportReason 
         Height          =   465
         Left            =   4500
         TabIndex        =   14
         Top             =   2550
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   820
         ForeColor       =   0
         MaxLength       =   40
         Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
      Begin UserControls.newText txtTransportWay 
         Height          =   465
         Left            =   4500
         TabIndex        =   15
         Top             =   3075
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   820
         ForeColor       =   0
         MaxLength       =   40
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
      Begin UserControls.newText txtLoadingSite 
         Height          =   465
         Left            =   4500
         TabIndex        =   16
         Top             =   3600
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   820
         ForeColor       =   0
         MaxLength       =   40
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
      Begin UserControls.newText txtDestinationSite 
         Height          =   465
         Left            =   4500
         TabIndex        =   17
         Top             =   4125
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   820
         ForeColor       =   0
         MaxLength       =   40
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
      Begin UserControls.newText txtInvoiceExtraRemarksA 
         Height          =   465
         Left            =   4500
         TabIndex        =   98
         Top             =   4650
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   820
         ForeColor       =   0
         MaxLength       =   255
         Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
      Begin UserControls.newText txtInvoiceExtraRemarksB 
         Height          =   465
         Left            =   4500
         TabIndex        =   99
         Top             =   5175
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   820
         ForeColor       =   0
         MaxLength       =   255
         Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
      Begin UserControls.newText txtUseNewInvoiceForm 
         Height          =   465
         Left            =   4500
         TabIndex        =   100
         Top             =   5700
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
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   6
         Left            =   5175
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   5700
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
         PicNormal       =   "UtilsSettings.frx":10CE
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Χρήση νέας φόρμας παραστατικού"
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
         Index           =   30
         Left            =   450
         TabIndex        =   104
         Top             =   5775
         Width           =   3615
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "1η γραμμή παρατηρήσεων"
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
         Left            =   450
         TabIndex        =   102
         Top             =   4725
         Width           =   3616
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "2η γραμμή παρατηρήσεων"
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
         Left            =   450
         TabIndex        =   101
         Top             =   5250
         Width           =   3616
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   20
         Left            =   4650
         Top             =   6150
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   6
         Left            =   8700
         Top             =   2775
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Τόπος φόρτωσης"
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
         Index           =   24
         Left            =   450
         TabIndex        =   49
         Top             =   3675
         Width           =   3615
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Τόπος προορισμού"
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
         Index           =   29
         Left            =   450
         TabIndex        =   60
         Top             =   4200
         Width           =   3615
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Σκοπός διακίνησης"
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
         TabIndex        =   59
         Top             =   2625
         Width           =   3615
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Τρόπος αποστολής"
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
         Index           =   23
         Left            =   450
         TabIndex        =   58
         Top             =   3150
         Width           =   3615
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   2
         Left            =   4050
         Top             =   450
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Στρογγυλοποίηση ποσών"
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
         TabIndex        =   55
         Top             =   1575
         Width           =   3615
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Εκτύπωση προηγούμενου - νέου υπολοίπου πελάτη"
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
         TabIndex        =   54
         Top             =   1050
         Width           =   3615
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Λεπτά στρογγυλοποίησης ποσών"
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
         Index           =   99
         Left            =   450
         TabIndex        =   52
         Top             =   2100
         Width           =   3615
         WordWrap        =   -1  'True
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
         BackColor       =   &H000080FF&
         Caption         =   "Εκτύπωση ώρας έκδοσης παραστατικού"
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
         TabIndex        =   51
         Top             =   525
         Width           =   3615
      End
   End
   Begin VB.Frame frmFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6690
      Index           =   4
      Left            =   6900
      TabIndex        =   92
      Top             =   8400
      Width           =   9165
      Begin UserControls.newText txtSender 
         Height          =   465
         Left            =   1950
         TabIndex        =   24
         Top             =   450
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   820
         ForeColor       =   0
         MaxLength       =   40
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
      Begin UserControls.newText txtServer 
         Height          =   465
         Left            =   1950
         TabIndex        =   25
         Top             =   975
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   820
         ForeColor       =   0
         MaxLength       =   40
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
      Begin UserControls.newText txtUserName 
         Height          =   465
         Left            =   1950
         TabIndex        =   26
         Top             =   1500
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   820
         ForeColor       =   0
         MaxLength       =   40
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
      Begin UserControls.newText txtPassword 
         Height          =   465
         Left            =   1950
         TabIndex        =   27
         Top             =   2025
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   820
         ForeColor       =   0
         MaxLength       =   40
         PasswordChar    =   "*"
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
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Κωδικός"
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
         TabIndex        =   96
         Top             =   2100
         Width           =   615
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   35
         Left            =   0
         Top             =   1050
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Ονομα χρήστη"
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
         TabIndex        =   95
         Top             =   1575
         Width           =   1065
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Διακομιστής"
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
         TabIndex        =   94
         Top             =   1050
         Width           =   915
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   33
         Left            =   1500
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
         Index           =   32
         Left            =   8700
         Top             =   1500
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Email"
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
         Left            =   450
         TabIndex        =   93
         Top             =   525
         Width           =   465
      End
   End
   Begin VB.Frame frmFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6690
      Index           =   3
      Left            =   2550
      TabIndex        =   62
      Top             =   1800
      Width           =   9165
      Begin VB.Frame frmEAFDSS 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   " Ε.Α.Φ.Δ.Σ.Σ. "
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1890
         Left            =   450
         TabIndex        =   63
         Top             =   450
         Width           =   8265
         Begin UserControls.newText txtEAFDSSProcessName 
            Height          =   465
            Left            =   3375
            TabIndex        =   19
            Top             =   1050
            Width           =   4440
            _ExtentX        =   7832
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   40
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
         Begin UserControls.newText txtEAFDSSCheckDescription 
            Height          =   465
            Left            =   3375
            TabIndex        =   18
            Top             =   525
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
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   4
            Left            =   4050
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   525
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
            PicNormal       =   "UtilsSettings.frx":1668
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   30
            Left            =   7800
            Top             =   825
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   540
            Index           =   21
            Left            =   3450
            Top             =   0
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Ονομα διαδικασίας"
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
            Index           =   25
            Left            =   450
            TabIndex        =   66
            Top             =   1125
            Width           =   2490
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   22
            Left            =   0
            Top             =   750
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Ελεγχος για φορτωμένη διαδικασία"
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
            TabIndex        =   65
            Top             =   600
            Width           =   2490
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   23
            Left            =   2925
            Top             =   825
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   390
            Index           =   24
            Left            =   4050
            Top             =   1500
            Visible         =   0   'False
            Width           =   465
         End
      End
      Begin UserControls.newDate mskClosedPeriod 
         Height          =   465
         Left            =   3375
         TabIndex        =   22
         Top             =   3525
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   820
         ForeColor       =   0
         Text            =   ""
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
      Begin UserControls.newFloat mskExtraChargesVATPercent 
         Height          =   465
         Left            =   3375
         TabIndex        =   21
         Top             =   3000
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   820
         Alignment       =   1
         ForeColor       =   0
         MaxLength       =   5
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
      Begin UserControls.newText txtTaxNoCheckDescription 
         Height          =   465
         Left            =   3375
         TabIndex        =   20
         Top             =   2475
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
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   5
         Left            =   4050
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   2475
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
         PicNormal       =   "UtilsSettings.frx":1C02
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newText txtBankAccountNo 
         Height          =   465
         Left            =   3375
         TabIndex        =   23
         Top             =   4050
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   820
         ForeColor       =   0
         MaxLength       =   50
         Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Height          =   540
         Index           =   7
         Left            =   4350
         Top             =   5625
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Τραπεζικός λογαριασμός"
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
         TabIndex        =   97
         Top             =   4125
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Ελεγχος Α.Φ.Μ. συναλλασόμενων"
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
         TabIndex        =   82
         Top             =   2550
         Width           =   2415
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   19
         Left            =   8700
         Top             =   1500
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   4
         Left            =   2925
         Top             =   3075
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Ποσοστό Φ.Π.Α. λοιπών χρεώσεων"
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
         Index           =   26
         Left            =   450
         TabIndex        =   68
         Top             =   3075
         Width           =   2490
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Κλεισμένη περίοδος έως"
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
         Index           =   28
         Left            =   450
         TabIndex        =   67
         Top             =   3600
         Width           =   2490
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   29
         Left            =   0
         Top             =   1050
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin VB.Frame frmFrame 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   6690
      Index           =   1
      Left            =   4050
      TabIndex        =   40
      Top             =   525
      Width           =   9165
      Begin VB.Frame Frame 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   " Επικεφαλίδες αναφορών "
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2940
         Index           =   1
         Left            =   450
         TabIndex        =   41
         Top             =   450
         Width           =   8265
         Begin UserControls.newText txtLine07 
            Height          =   465
            Left            =   1650
            TabIndex        =   6
            Top             =   525
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin UserControls.newText txtLine08 
            Height          =   465
            Left            =   1650
            TabIndex        =   7
            Top             =   1050
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin UserControls.newText txtLine09 
            Height          =   465
            Left            =   1650
            TabIndex        =   8
            Top             =   1575
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin UserControls.newText txtLine10 
            Height          =   465
            Left            =   1650
            TabIndex        =   9
            Top             =   2100
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
            BackColor       =   &H000080FF&
            Caption         =   "1η Γραμμή"
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
            TabIndex        =   45
            Top             =   600
            Width           =   750
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "2η Γραμμή"
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
            Left            =   450
            TabIndex        =   44
            Top             =   1125
            Width           =   750
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "3η Γραμμή"
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
            TabIndex        =   43
            Top             =   1650
            Width           =   750
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "4η Γραμμή"
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
            TabIndex        =   42
            Top             =   2175
            Width           =   750
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   540
            Index           =   15
            Left            =   2550
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
            Index           =   16
            Left            =   3225
            Top             =   2550
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   17
            Left            =   0
            Top             =   1125
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   540
            Index           =   18
            Left            =   7800
            Top             =   1200
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   31
            Left            =   1200
            Top             =   1050
            Visible         =   0   'False
            Width           =   465
         End
      End
      Begin UserControls.newText txtPreviewReportsDescription 
         Height          =   465
         Left            =   2850
         TabIndex        =   46
         Top             =   3525
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
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   0
         Left            =   3525
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   3525
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
         PicNormal       =   "UtilsSettings.frx":219C
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   14
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Προεπισκόπηση αναφορών"
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
         Index           =   27
         Left            =   450
         TabIndex        =   48
         Top             =   3600
         Width           =   1965
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   34
         Left            =   2400
         Top             =   3525
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin VB.Frame frmInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   15750
      TabIndex        =   69
      Top             =   2475
      Width           =   4515
      Begin VB.TextBox txtUseNewInvoiceFormID 
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
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   2850
         Width           =   780
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
         TabIndex        =   105
         TabStop         =   0   'False
         Text            =   "Settings.UseNewInvoiceForm"
         Top             =   2850
         Width           =   3540
      End
      Begin VB.TextBox Text8 
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
         TabIndex        =   85
         TabStop         =   0   'False
         Text            =   "Settings.TaxNoCheckID"
         Top             =   2475
         Width           =   3540
      End
      Begin VB.TextBox txtTaxNoCheckID 
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
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   2475
         Width           =   780
      End
      Begin VB.TextBox txtEAFDSSCheckID 
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
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   2100
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
         TabIndex        =   80
         TabStop         =   0   'False
         Text            =   "Settings.EAFDSSCheckID"
         Top             =   2100
         Width           =   3540
      End
      Begin VB.TextBox txtRoundSalesID 
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
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   1725
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
         TabIndex        =   78
         TabStop         =   0   'False
         Text            =   "Settings.RoundSalesID"
         Top             =   1725
         Width           =   3540
      End
      Begin VB.TextBox txtPrintBalanceID 
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
         TabIndex        =   77
         TabStop         =   0   'False
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
         TabIndex        =   76
         TabStop         =   0   'False
         Text            =   "Settings.PrintBalanceID"
         Top             =   1350
         Width           =   3540
      End
      Begin VB.TextBox txtPrintHourID 
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
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   975
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
         TabIndex        =   74
         TabStop         =   0   'False
         Text            =   "Settings.PrintHourID"
         Top             =   975
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
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   225
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
         TabIndex        =   72
         TabStop         =   0   'False
         Text            =   "Settings.ID"
         Top             =   225
         Width           =   3540
      End
      Begin VB.TextBox txtPreviewReportsID 
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
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   600
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
         TabIndex        =   70
         TabStop         =   0   'False
         Text            =   "Settings.PreviewReportsID"
         Top             =   600
         Width           =   3540
      End
   End
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   0
      TabIndex        =   86
      Top             =   8325
      Width           =   6090
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   87
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
         Index           =   1
         Left            =   1650
         TabIndex        =   88
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
         Index           =   3
         Left            =   4500
         TabIndex        =   89
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
         TabIndex        =   90
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
   Begin Dacara_dcButton.dcButton btnPanel 
      Height          =   990
      Index           =   0
      Left            =   450
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1125
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1746
      BackColor       =   12640511
      ButtonShape     =   3
      ButtonStyle     =   2
      Caption         =   "Παραστατικά πωλήσεων"
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
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2175
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1746
      BackColor       =   12640511
      ButtonShape     =   3
      ButtonStyle     =   2
      Caption         =   "Αναφορές"
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
   Begin Dacara_dcButton.dcButton btnPanel 
      Height          =   990
      Index           =   2
      Left            =   450
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3225
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1746
      BackColor       =   12640511
      ButtonShape     =   3
      ButtonStyle     =   2
      Caption         =   "Πωλήσεις"
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
   Begin Dacara_dcButton.dcButton btnPanel 
      Height          =   990
      Index           =   3
      Left            =   450
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   4275
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1746
      BackColor       =   12640511
      ButtonShape     =   3
      ButtonStyle     =   2
      Caption         =   "Καθολικές ρυθμίσεις"
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
   Begin Dacara_dcButton.dcButton btnPanel 
      Height          =   990
      Index           =   4
      Left            =   450
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   5325
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1746
      BackColor       =   12640511
      ButtonShape     =   3
      ButtonStyle     =   2
      Caption         =   "Ρυθμίσεις email"
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
   Begin VB.Shape shpWedge 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   540
      Index           =   25
      Left            =   4950
      Top             =   7800
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBridge 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1005
      Index           =   4
      Left            =   450
      Top             =   5325
      Width           =   3090
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   5175
      Top             =   9000
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   11025
      Top             =   2250
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBridge 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1005
      Index           =   3
      Left            =   450
      Top             =   4275
      Width           =   3090
   End
   Begin VB.Shape shpBridge 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1005
      Index           =   2
      Left            =   450
      Top             =   3225
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
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Παραμετροποίηση"
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
      Left            =   225
      TabIndex        =   28
      Top             =   75
      Width           =   4305
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   3
      Left            =   1875
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
      Top             =   1500
      Visible         =   0   'False
      Width           =   465
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
Attribute VB_Name = "UtilsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnStatus As Boolean

Private Function PositionPanels()

    Dim intLoop As Integer
    
    For intLoop = 0 To 4
        frmFrame(intLoop).Visible = False
    Next intLoop
        
    For intLoop = 0 To 4
        btnPanel(intLoop).Enabled = True
        shpBridge(intLoop).Visible = False
        With frmFrame(intLoop)
            .Left = 1875
            .Top = 1125
            .BackColor = &HE0E0E0
        End With
    Next intLoop
    
    btnPanel(0).Enabled = False
    frmFrame(0).Visible = True
    shpBridge(0).Visible = True

End Function

Private Function LoadSettings()

    Dim rsParameters As Recordset
    
    Set rsParameters = CommonDB.OpenRecordset("Parameters")
    With rsParameters
        .MoveFirst
        txtID.text = !ID
        txtLine01.text = !Line01
        txtLine02.text = !Line02
        txtLine03.text = !Line03
        txtLine04.text = !Line04
        txtLine05.text = !Line05
        txtLine06.text = !Line06
        txtLine07.text = !Line07
        txtLine08.text = !Line08
        txtLine09.text = !Line09
        txtLine10.text = !Line10
        txtPreviewReportsID = !PreviewReportsID
        txtPrintHourID.text = !PrintHourID
        txtPrintBalanceID.text = !PrintBalanceID
        txtRoundSalesID.text = !RoundSalesID
        mskRoundSalesCents.text = !RoundSalesCents
        txtTransportReason.text = !TransportReason
        txtTransportWay.text = !TransportWay
        txtLoadingSite.text = !LoadingSite
        txtDestinationSite.text = !DestinationSite
        txtEAFDSSCheckID.text = !EAFDSSCheckID
        txtEAFDSSProcessName.text = !EAFDSSProcessName
        txtTaxNoCheckID.text = !TaxNoCheckID
        mskExtraChargesVATPercent.text = !ExtraChargesVATPercent
        mskClosedPeriod.text = !ClosedPeriod
        txtBankAccountNo.text = !BankAccountNo
        txtUseNewInvoiceFormID.text = !UseNewInvoiceForm
        .Close
    End With

End Function

Private Function SaveRecord()

    If Not ValidateFields Then Exit Function
    
    If MainSaveRecord("CommonDB", "Settings", False, "Settings", "ID", txtID.text, txtLine01.text, txtLine02.text, txtLine03.text, txtLine04.text, txtLine05.text, txtLine06.text, txtLine07.text, txtLine08.text, txtLine09.text, txtLine10.text, txtPreviewReportsID.text, txtPrintHourID.text, txtPrintBalanceID.text, txtRoundSalesID.text, mskRoundSalesCents.text, txtTransportReason.text, txtTransportWay.text, txtLoadingSite.text, txtDestinationSite.text, txtEAFDSSCheckID.text, txtEAFDSSProcessName.text, txtTaxNoCheckID.text, mskExtraChargesVATPercent.text, mskClosedPeriod.text, txtBankAccountNo.text, txtInvoiceExtraRemarksA.text, txtInvoiceExtraRemarksB.text, txtUseNewInvoiceFormID.text, txtSender.text, txtServer.text, txtUserName.text, txtPassword.text) <> 0 Then
        btnPanel_Click 0
        blnStatus = True
        DisableFields txtLine01, txtLine02, txtLine03, txtLine04, txtLine05, txtLine06, txtLine07, txtLine08, txtLine09, txtLine10, txtPreviewReportsDescription, txtPrintHourDescription, txtPrintBalanceDescription, txtRoundSalesDescription, mskRoundSalesCents, txtTransportReason, txtTransportWay, txtLoadingSite, txtDestinationSite, txtEAFDSSCheckDescription, txtTaxNoCheckDescription, txtEAFDSSProcessName, mskExtraChargesVATPercent, mskClosedPeriod, txtBankAccountNo, txtSender, txtServer, txtUserName, txtPassword, txtInvoiceExtraRemarksA, txtInvoiceExtraRemarksB, txtUseNewInvoiceForm
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6)
        UpdateButtons Me, 3, 1, 0, 0, 1
        If MyMsgBox(1, strAppTitle, strMessages(22), 1) Then
        End If
    Else
        DisplayErrorMessage True, strMessages(26)
    End If
    
End Function

Private Sub btnPanel_Click(Index As Integer)

    Dim intLoop As Integer
    
    For intLoop = 0 To 4
        btnPanel(intLoop).Enabled = True
        frmFrame(intLoop).Visible = False
        shpBridge(intLoop).Visible = False
    Next intLoop
    
    btnPanel(Index).Enabled = False
    frmFrame(Index).Visible = True
    shpBridge(Index).Visible = True
    
    Select Case Index
        'Επικεφαλίδες παραστατικών
        Case 0
            If cmdButton(1).Enabled Then
                If txtLine01.Enabled Then txtLine01.SetFocus
            End If
        'Επικεφαλίδες αναφορών
        Case 1
            If cmdButton(1).Enabled Then
                If txtLine07.Enabled Then txtLine07.SetFocus
            End If
        'Πωλήσεις
        Case 2
            If cmdButton(1).Enabled Then
                If txtPrintHourDescription.Enabled Then txtPrintHourDescription.SetFocus
            End If
        'Καθολικές ρυθμίσεις
        Case 3
            If cmdButton(1).Enabled Then
                If txtEAFDSSCheckDescription.Enabled Then txtEAFDSSCheckDescription.SetFocus
            End If
        'Email
        Case 4
            If cmdButton(1).Enabled Then
                If txtSender.Enabled Then txtSender.SetFocus
            End If
    End Select

End Sub

Private Function GotoNextPanel(formName, panelCount)

    Dim intLoop As Integer
    
    For intLoop = 0 To panelCount - 1
    
        If Not formName.btnPanel(intLoop).Enabled Then
            If intLoop + 1 <= formName.btnPanel.Count - 1 Then
                If formName.btnPanel(intLoop + 1).Enabled Then
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

Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
        Case 0
            EditRecord
        Case 1
            SaveRecord
        Case 2
            AbortProcedure False
        Case 3
            AbortProcedure True
    End Select

End Sub

Private Function EditRecord()

    Dim intLoop As Integer
    
    blnStatus = False
    
    EnableFields txtLine01, txtLine02, txtLine03, txtLine04, txtLine05, txtLine06, txtLine07, txtLine08, txtLine09, txtLine10, txtPreviewReportsDescription, txtPrintHourDescription, txtPrintBalanceDescription, txtRoundSalesDescription, mskRoundSalesCents, txtTransportReason, txtTransportWay, txtLoadingSite, txtDestinationSite, txtEAFDSSCheckDescription, txtEAFDSSProcessName, txtTaxNoCheckDescription, mskExtraChargesVATPercent, mskClosedPeriod, txtBankAccountNo, txtInvoiceExtraRemarksA, txtInvoiceExtraRemarksB, txtSender, txtServer, txtUserName, txtPassword, txtUseNewInvoiceForm
    EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6)
    
    UpdateButtons Me, 3, 0, 1, 1, 0
    
    For intLoop = 0 To btnPanel.Count - 1
        If Not btnPanel(intLoop).Enabled Then btnPanel_Click intLoop
    Next intLoop

End Function

Private Function AbortProcedure(blnStatus)
    
    If Not blnStatus Then
        If MyMsgBox(3, strAppTitle, strMessages(3), 2) Then
            btnPanel_Click 0
            blnStatus = False
            DisableFields txtLine01, txtLine02, txtLine03, txtLine04, txtLine05, txtLine06, txtLine07, txtLine08, txtLine09, txtLine10, txtPreviewReportsDescription, txtPrintHourDescription, txtPrintBalanceDescription, txtRoundSalesDescription, mskRoundSalesCents, txtTransportReason, txtTransportWay, txtLoadingSite, txtDestinationSite, txtEAFDSSCheckDescription, txtEAFDSSProcessName, txtTaxNoCheckDescription, mskExtraChargesVATPercent, mskClosedPeriod, txtBankAccountNo, txtInvoiceExtraRemarksA, txtInvoiceExtraRemarksB, txtSender, txtServer, txtUserName, txtPassword, txtUseNewInvoiceForm
            DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6)
            UpdateButtons Me, 3, 1, 0, 0, 1
        End If
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Sub cmdIndex_Click(Index As Integer)

    'Local μεταβλητές
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case Index
        Case 0
            'Προεπισκόπηση αναφορών
            If txtPreviewReportsDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtPreviewReportsDescription.text, "YesOrNo", "YesNoDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtPreviewReportsID.text = tmpTableData.strCode
            txtPreviewReportsDescription.text = tmpTableData.strOneField
        Case 1
            'Εκτύπωση ώρας έκδοσης παραστατικού πωλήσεων
            If txtPrintHourDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtPrintHourDescription.text, "YesOrNo", "YesNoDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtPrintHourID.text = tmpTableData.strCode
            txtPrintHourDescription.text = tmpTableData.strOneField
        Case 2
            'Εκτύπωση προηγούμενου - νέου υπολοίπου πελάτη
            If txtPrintBalanceDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtPrintBalanceDescription.text, "YesOrNo", "YesNoDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtPrintBalanceID.text = tmpTableData.strCode
            txtPrintBalanceDescription.text = tmpTableData.strOneField
        Case 3
            'Στρογγυλοποίηση ποσών πωλήσεων
            If txtRoundSalesDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtRoundSalesDescription.text, "YesOrNo", "YesNoDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtRoundSalesID.text = tmpTableData.strCode
            txtRoundSalesDescription.text = tmpTableData.strOneField
        Case 4
            'Ελεγχος για φορτωμένη διαδικασία ΕΑΦΔΣΣ
            If txtEAFDSSCheckDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtEAFDSSCheckDescription.text, "YesOrNo", "YesNoDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtEAFDSSCheckID.text = tmpTableData.strCode
            txtEAFDSSCheckDescription.text = tmpTableData.strOneField
        Case 5
            'Ελεγχος Α.Φ.Μ. συναλλασόμενων
            If txtTaxNoCheckDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtTaxNoCheckDescription.text, "YesOrNo", "YesNoDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtTaxNoCheckID.text = tmpTableData.strCode
            txtTaxNoCheckDescription.text = tmpTableData.strOneField
        Case 6
            'Χρήση νέας φόρμας παραστατικού
            If txtUseNewInvoiceForm.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtUseNewInvoiceForm.text, "YesOrNo", "YesNoDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtUseNewInvoiceFormID.text = tmpTableData.strCode
            txtUseNewInvoiceForm.text = tmpTableData.strOneField
    End Select

End Sub

Private Sub Form_Activate()

    Dim tmpRecordset As Recordset

    If Me.Tag = "True" Then
        Me.Tag = "False"
        Me.Refresh
        If MainSeekRecord("CommonDB", "Settings", "ID", 1, True, txtID, txtLine01, txtLine02, txtLine03, txtLine04, txtLine05, txtLine06, txtLine07, txtLine08, txtLine09, txtLine10, txtPreviewReportsID, txtPrintHourID, txtPrintBalanceID, txtRoundSalesID, mskRoundSalesCents, txtTransportReason, txtTransportWay, txtLoadingSite, txtDestinationSite, txtEAFDSSCheckID, txtEAFDSSProcessName, txtTaxNoCheckID, mskExtraChargesVATPercent, mskClosedPeriod, txtBankAccountNo, txtInvoiceExtraRemarksA, txtInvoiceExtraRemarksB, txtUseNewInvoiceFormID, txtSender, txtServer, txtUserName, txtPassword) Then
            'Προεπισκόπηση αναφορών
            Set tmpRecordset = CheckForMatch("CommonDB", txtPreviewReportsID.text, "YesOrNo", "YesNoID", "Numeric", 0, 1)
            txtPreviewReportsID.text = tmpRecordset.Fields(0)
            txtPreviewReportsDescription.text = tmpRecordset.Fields(1)
            'Εκτύπωση ώρας έκδοσης παραστατικού πωλήσεων
            Set tmpRecordset = CheckForMatch("CommonDB", txtPrintHourID.text, "YesOrNo", "YesNoID", "Numeric", 0, 1)
            txtPrintHourID.text = tmpRecordset.Fields(0)
            txtPrintHourDescription.text = tmpRecordset.Fields(1)
            'Εκτύπωση προηγούμενου - νέου υπόλοιπου πελάτη
            Set tmpRecordset = CheckForMatch("CommonDB", txtPrintBalanceID.text, "YesOrNo", "YesNoID", "Numeric", 0, 1)
            txtPrintBalanceID.text = tmpRecordset.Fields(0)
            txtPrintBalanceDescription.text = tmpRecordset.Fields(1)
            'Στρογγυλοποίηση ποσών πωλήσεων
            Set tmpRecordset = CheckForMatch("CommonDB", txtRoundSalesID.text, "YesOrNo", "YesNoID", "Numeric", 0, 1)
            txtRoundSalesID.text = tmpRecordset.Fields(0)
            txtRoundSalesDescription.text = tmpRecordset.Fields(1)
            'Ελεγχος για φορτωμένη διαδικασία ΕΑΦΔΣΣ
            Set tmpRecordset = CheckForMatch("CommonDB", txtEAFDSSCheckID.text, "YesOrNo", "YesNoID", "Numeric", 0, 1)
            txtEAFDSSCheckID.text = tmpRecordset.Fields(0)
            txtEAFDSSCheckDescription.text = tmpRecordset.Fields(1)
            'Ελεγχος για Α.Φ.Μ. συναλλασόμενων
            Set tmpRecordset = CheckForMatch("CommonDB", txtTaxNoCheckID.text, "YesOrNo", "YesNoID", "Numeric", 0, 1)
            txtTaxNoCheckID.text = tmpRecordset.Fields(0)
            txtTaxNoCheckDescription.text = tmpRecordset.Fields(1)
            'Χρήση νέας φόρμας παραστατικού
            Set tmpRecordset = CheckForMatch("CommonDB", txtUseNewInvoiceFormID.text, "YesOrNo", "YesNoID", "Numeric", 0, 1)
            txtUseNewInvoiceFormID.text = tmpRecordset.Fields(0)
            txtUseNewInvoiceForm.text = tmpRecordset.Fields(1)
            '
            UpdateButtons Me, 3, 1, 0, 0, 1
        End If
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
        Case vbKeyE And CtrlDown = 4 And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyF10 And cmdButton(1).Enabled, vbKeyS And CtrlDown = 4 And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyEscape
            If cmdButton(2).Enabled Then cmdButton_Click 2: Exit Function
            If cmdButton(3).Enabled Then cmdButton_Click 3
        Case vbKeyPageUp
            GotoPreviousPanel Me, btnPanel.Count
        Case vbKeyPageDown
            GotoNextPanel Me, btnPanel.Count
        Case vbKeyF12
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    PositionPanels
    PositionControls Me, False: ColorizeControls Me
    ClearFields txtLine01, txtLine02, txtLine03, txtLine04, txtLine05, txtLine06, txtLine07, txtLine08, txtLine09, txtLine10, txtPreviewReportsDescription, txtPrintHourDescription, txtPrintBalanceDescription, txtRoundSalesDescription, mskRoundSalesCents, txtTransportReason, txtTransportWay, txtLoadingSite, txtDestinationSite, txtEAFDSSCheckDescription, txtEAFDSSProcessName, txtTaxNoCheckDescription, mskExtraChargesVATPercent, mskClosedPeriod, txtBankAccountNo, txtSender, txtServer, txtUserName, txtPassword, txtUseNewInvoiceForm, txtInvoiceExtraRemarksA, txtInvoiceExtraRemarksB
    DisableFields txtLine01, txtLine02, txtLine03, txtLine04, txtLine05, txtLine06, txtLine07, txtLine08, txtLine09, txtLine10, txtPreviewReportsDescription, txtPrintHourDescription, txtPrintBalanceDescription, txtRoundSalesDescription, mskRoundSalesCents, txtTransportReason, txtTransportWay, txtLoadingSite, txtDestinationSite, txtEAFDSSCheckDescription, txtEAFDSSProcessName, txtTaxNoCheckDescription, mskExtraChargesVATPercent, mskClosedPeriod, txtBankAccountNo, txtSender, txtServer, txtUserName, txtPassword, txtUseNewInvoiceForm, txtInvoiceExtraRemarksA, txtInvoiceExtraRemarksB
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6)
    UpdateButtons Me, 3, 1, 0, 0, 1
    
End Sub

Private Function ValidateFields()

    ValidateFields = False
    
    'Προεπισκόπηση αναφορών
    If DisplayMessage(1, 4, 1, "", txtPreviewReportsID.text) Then
        btnPanel_Click 1
        txtPreviewReportsDescription.SetFocus
        Exit Function
    End If
    
    'Εκτύπωση ώρας έκδοσης παραστατικού
    If DisplayMessage(1, 4, 1, "", txtPrintHourID.text) Then
        btnPanel_Click 2
        txtPrintHourDescription.SetFocus
        Exit Function
    End If
    
    'Εκτύπωση προηγούμενου - νέου υπολοίπου πελάτη
    If DisplayMessage(1, 4, 1, "", txtPrintBalanceID.text) Then
        btnPanel_Click 2
        txtPrintBalanceDescription.SetFocus
        Exit Function
    End If
    
    'Στρογγυλοποίηση ποσών
    If DisplayMessage(1, 4, 1, "", txtRoundSalesID.text) Then
        btnPanel_Click 2
        txtRoundSalesDescription.SetFocus
        Exit Function
    End If
    
    'Λεπτά στρογγυλοποίησης ποσών
    If DisplayMessage(1, 4, 1, "", mskRoundSalesCents.text) Then
        btnPanel_Click 2
        mskRoundSalesCents.SetFocus
        Exit Function
    End If
    
    'Χρήση νέας φόρμας παραστατικού
    If DisplayMessage(1, 4, 1, "", txtUseNewInvoiceFormID.text) Then
        btnPanel_Click 2
        txtUseNewInvoiceForm.SetFocus
        Exit Function
    End If
    
    'Ελεγχος για φορτωμένη διαδικασία ΕΑΦΔΣΣ
    If DisplayMessage(1, 4, 1, "", txtEAFDSSCheckID.text) Then
        btnPanel_Click 3
        txtEAFDSSCheckDescription.SetFocus
        Exit Function
    End If
    
    'Ελεγχος για Α.Φ.Μ. συναλλασόμενων
    If DisplayMessage(1, 4, 1, "", txtTaxNoCheckID.text) Then
        btnPanel_Click 3
        txtTaxNoCheckDescription.SetFocus
        Exit Function
    End If
    
    'Ποσοστό Φ.Π.Α. λοιπών χρεώσεων
    If DisplayMessage(1, 4, 1, "", mskExtraChargesVATPercent.text) Then
        btnPanel_Click 3
        mskExtraChargesVATPercent.SetFocus
        Exit Function
    End If
    
    'Κλεισμένη περίοδος έως
    If mskClosedPeriod.text = "" Or Not IsDate(mskClosedPeriod.text) Then
        If MyMsgBox(4, strAppTitle, strMessages(1), 1) Then
        End If
        btnPanel_Click 3
        mskClosedPeriod.SetFocus
        Exit Function
    End If
    
    ValidateFields = True

End Function

Private Sub txtEAFDSSCheckDescription_Change()

    If txtEAFDSSCheckDescription.text = "" Then ClearFields txtEAFDSSCheckID

End Sub

Private Sub txtEAFDSSCheckDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 4

End Sub

Private Sub txtEAFDSSCheckDescription_Validate(Cancel As Boolean)

    If txtEAFDSSCheckID.text = "" And txtEAFDSSCheckDescription.text <> "" Then cmdIndex_Click 4: If txtEAFDSSCheckID.text = "" Then Cancel = True

End Sub

Private Sub txtPreviewReportsDescription_Change()

    If txtPreviewReportsDescription.text = "" Then ClearFields txtPreviewReportsID

End Sub

Private Sub txtPreviewReportsDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub

Private Sub txtPreviewReportsDescription_Validate(Cancel As Boolean)

    If txtPreviewReportsID.text = "" And txtPreviewReportsDescription.text <> "" Then cmdIndex_Click 0: If txtPreviewReportsID.text = "" Then Cancel = True

End Sub

Private Sub txtPrintBalanceDescription_Change()

    If txtPrintBalanceDescription.text = "" Then ClearFields txtPrintBalanceID

End Sub

Private Sub txtPrintBalanceDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2

End Sub

Private Sub txtPrintBalanceDescription_Validate(Cancel As Boolean)

    If txtPrintBalanceID.text = "" And txtPrintBalanceDescription.text <> "" Then cmdIndex_Click 2: If txtPrintBalanceID.text = "" Then Cancel = True

End Sub

Private Sub txtPrintHourDescription_Change()

    If txtPrintHourDescription.text = "" Then ClearFields txtPrintHourID

End Sub

Private Sub txtPrintHourDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1

End Sub

Private Sub txtPrintHourDescription_Validate(Cancel As Boolean)

    If txtPrintHourID.text = "" And txtPrintHourDescription.text <> "" Then cmdIndex_Click 1: If txtPrintHourID.text = "" Then Cancel = True

End Sub

Private Sub txtRoundSalesDescription_Change()

    If txtRoundSalesDescription.text = "" Then ClearFields txtRoundSalesID

End Sub

Private Sub txtRoundSalesDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 3

End Sub

Private Sub txtRoundSalesDescription_Validate(Cancel As Boolean)

    If txtRoundSalesID.text = "" And txtRoundSalesDescription.text <> "" Then cmdIndex_Click 3: If txtRoundSalesID.text = "" Then Cancel = True

End Sub

Private Sub txtTaxNoCheckDescription_Change()
    
    If txtTaxNoCheckDescription.text = "" Then ClearFields txtTaxNoCheckID

End Sub

Private Sub txtTaxNoCheckDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 5

End Sub

Private Sub txtTaxNoCheckDescription_Validate(Cancel As Boolean)

    If txtTaxNoCheckID.text = "" And txtTaxNoCheckDescription.text <> "" Then cmdIndex_Click 5: If txtTaxNoCheckID.text = "" Then Cancel = True

End Sub

Private Sub txtUseNewInvoiceForm_Change()

    If txtUseNewInvoiceForm.text = "" Then ClearFields txtUseNewInvoiceFormID

End Sub


Private Sub txtUseNewInvoiceForm_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 6

End Sub


Private Sub txtUseNewInvoiceForm_Validate(Cancel As Boolean)

    If txtUseNewInvoiceFormID.text = "" And txtUseNewInvoiceForm.text <> "" Then cmdIndex_Click 6: If txtUseNewInvoiceFormID.text = "" Then Cancel = True

End Sub


