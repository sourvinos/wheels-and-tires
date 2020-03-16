VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form CommonPendingInvoices 
   Appearance      =   0  'Flat
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
      Left            =   9525
      TabIndex        =   13
      Top             =   7650
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "CommonPendingInvoices.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "CommonPendingInvoices.frx":001C
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
         TabIndex        =   15
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
         TabIndex        =   28
         Top             =   8850
         Width           =   11790
         Begin GurhanButtonOCX.GurhanButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   29
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
            TabIndex        =   30
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
            Index           =   7
            Left            =   10200
            TabIndex        =   31
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
            Index           =   4
            Left            =   5925
            TabIndex        =   32
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
            Index           =   6
            Left            =   8775
            TabIndex        =   33
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
            Index           =   5
            Left            =   7350
            TabIndex        =   34
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
         Begin GurhanButtonOCX.GurhanButton cmdButton 
            Height          =   690
            Index           =   2
            Left            =   3075
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            Caption         =   "Τιμολόγηση"
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
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            Caption         =   "Γρήγορη τιμολόγηση"
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
         Height          =   2940
         Left            =   9450
         TabIndex        =   19
         Tag             =   "Hidden"
         Top             =   4575
         Visible         =   0   'False
         Width           =   4515
         Begin VB.TextBox txtInitialRefersTo 
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
            TabIndex        =   49
            TabStop         =   0   'False
            Text            =   "5"
            Top             =   1575
            Width           =   780
         End
         Begin VB.TextBox Text2 
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
            TabIndex        =   48
            TabStop         =   0   'False
            Text            =   "InitialRefersTo"
            Top             =   1575
            Width           =   3540
         End
         Begin VB.TextBox Text5 
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
            TabIndex        =   47
            TabStop         =   0   'False
            Text            =   "IsTriangular"
            Top             =   1950
            Width           =   3540
         End
         Begin VB.TextBox txtTriangularID 
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
            TabIndex        =   46
            TabStop         =   0   'False
            Text            =   "6"
            Top             =   1950
            Width           =   780
         End
         Begin VB.TextBox txtDeliveryPointID 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
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
            TabIndex        =   45
            TabStop         =   0   'False
            Text            =   "2"
            Top             =   450
            Width           =   780
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
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
            Text            =   "DeliveryPoints.DeliveryPointID"
            Top             =   450
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
            TabIndex        =   38
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
            TabIndex        =   37
            TabStop         =   0   'False
            Text            =   "RefersTo"
            Top             =   1200
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
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "3"
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
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   "Table"
            Top             =   825
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
            TabIndex        =   26
            TabStop         =   0   'False
            Text            =   "Persons.PersonID"
            Top             =   75
            Width           =   3540
         End
         Begin VB.TextBox txtPersonID 
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
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   "1"
            Top             =   75
            Width           =   780
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   2325
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   2296
            Images          =   "CommonPendingInvoices.frx":0038
            Version         =   131072
            KeyCount        =   2
            Keys            =   "ÿ"
         End
      End
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   3840
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   4875
         Width           =   9240
         Begin VB.CheckBox chkCriteriaItemAnalysis 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Caption         =   "Ανάλυση ειδών"
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
            Height          =   315
            Left            =   2175
            TabIndex        =   7
            Top             =   2925
            Value           =   1  'Checked
            Width           =   4065
         End
         Begin UserControls.newText txtPersonDescription 
            Height          =   465
            Left            =   2175
            TabIndex        =   5
            Top             =   1875
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
         Begin UserControls.newDate mskIssueFrom 
            Height          =   465
            Left            =   2175
            TabIndex        =   1
            Top             =   825
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   820
            ForeColor       =   0
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
         Begin UserControls.newDate mskIssueTo 
            Height          =   465
            Left            =   3750
            TabIndex        =   2
            Top             =   825
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   820
            ForeColor       =   0
            Text            =   "31/12/2017"
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
            Left            =   8400
            TabIndex        =   24
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
            PicNormal       =   "CommonPendingInvoices.frx":0950
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newDate mskInFrom 
            Height          =   465
            Left            =   2175
            TabIndex        =   3
            Top             =   1350
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   820
            ForeColor       =   0
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
         Begin UserControls.newDate mskInTo 
            Height          =   465
            Left            =   3750
            TabIndex        =   4
            Top             =   1350
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   820
            ForeColor       =   0
            Text            =   "31/12/2017"
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
         Begin UserControls.newText txtDeliveryPointDescription 
            Height          =   465
            Left            =   2175
            TabIndex        =   6
            Top             =   2400
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   820
            ForeColor       =   0
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
            Index           =   1
            Left            =   7200
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   2400
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
            PicNormal       =   "CommonPendingInvoices.frx":0EEA
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000000C0&
            Caption         =   "Τόπος παραλαβής"
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
            Top             =   2475
            Width           =   1290
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000000C0&
            Caption         =   "Καταχώρηση"
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
            TabIndex        =   39
            Top             =   1425
            Width           =   1290
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000000C0&
            Caption         =   "Συναλλασόμενος"
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
            TabIndex        =   23
            Top             =   1950
            Width           =   1290
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
            Left            =   4575
            TabIndex        =   20
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
            Left            =   1725
            Top             =   825
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
            Left            =   8775
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
            TabIndex        =   18
            Top             =   3375
            Width           =   9240
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
            TabIndex        =   16
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
            TabIndex        =   9
            Top             =   900
            Width           =   1290
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
            TabIndex        =   17
            Top             =   0
            Width           =   9240
         End
      End
      Begin iGrid300_10Tec.iGrid grdCommonPendingInvoices 
         Height          =   7290
         Left            =   75
         TabIndex        =   10
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
         TabIndex        =   27
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
         TabIndex        =   22
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
         ForeColor       =   &H00C0C000&
         Height          =   315
         Left            =   75
         TabIndex        =   21
         Top             =   1125
         Width           =   2565
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Εκκρεμή δελτία αποστολής"
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
         Height          =   765
         Left            =   75
         TabIndex        =   12
         Top             =   75
         Width           =   6315
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
         TabIndex        =   11
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
Attribute VB_Name = "CommonPendingInvoices"
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
Dim intLastNo As Integer

Private Function QuickTransformInvoices()

    On Error GoTo ErrTrap
    
    Dim lngRow As Long
    Dim rsInvoices As Recordset
    
    If Not ValidateSelectedLines Then Exit Function
    
    BeginTrans
    
    Set rsInvoices = CommonDB.OpenRecordset("Invoices")
    rsInvoices.Index = "TrnID"
    
    With grdCommonPendingInvoices
        For lngRow = 1 To .RowCount
            If .CellIcon(lngRow, "Selected") > 0 Then
                rsInvoices.Seek "=", Val(.CellText(lngRow, "InvoiceTrnID"))
                If Not rsInvoices.NoMatch Then
                    rsInvoices.Edit
                    rsInvoices!InvoiceIsInvoiced = 2
                    rsInvoices.Update
                End If
            End If
        Next lngRow
    End With
    
    CommitTrans
    
    rsInvoices.Close
    
    DisplayMessage 10, 1, 1, ""
    
    Exit Function
    
ErrTrap:
    Rollback
    DisplayErrorMessage True, Err.Description

End Function

Private Function SeekAndEditRecord(myInvoiceTrnID, myWindowTitle, myNextWindowTitle, myTable, myRefersTo, myOppositeTable, myOppositeRefersTo)
    
    Dim blnFound As Boolean
    
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

End Function

Private Function FindRecordsAndPopulateGrid()

    If RefreshList > 0 Then
        UpdateRecordCount lblRecordCount, lngRowCount
        UpdateCriteriaLabels mskIssueFrom.text, mskIssueTo.text, mskInFrom.text, mskInTo.text, txtPersonDescription.text
        EnableGrid grdCommonPendingInvoices, False
        HighlightRow grdCommonPendingInvoices, 1, "", True
        UpdateButtons Me, 7, 0, 1, 1, 1, 1, 1, 1, 0
    Else
        UpdateButtons Me, 7, 1, 0, 0, 0, 0, 0, 0, 1
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
        mskIssueFrom.SetFocus
    End If
    
End Function

Private Function AddGridRowWithTotals(myGrid As iGrid, myOnlyQty, myMessageColumn, myPrintableLineOrNot, myMessage, myBalance, mySums, myColumnCount, myHowManyBlankLinesBefore, myHowManyBlankLinesAfter, ParamArray myColumns() As Variant)

    Dim intLoop As Integer
    Dim lngRow As Long
    
    myGrid.Redraw = False
    
    myGrid.AddRow , , , , , , , myHowManyBlankLinesBefore
    lngRow = myGrid.RowCount
    
    myGrid.CellValue(lngRow, myMessageColumn) = myMessage
    
    For intLoop = 0 To myColumnCount - IIf(myBalance, 1, 1)
        myGrid.CellValue(lngRow, myColumns(intLoop)) = IIf(myOnlyQty = 0, mySums(intLoop), 0)
    Next intLoop
    
    If Not myBalance Then
        myGrid.CellValue(lngRow, "Balance") = mySums(0) - mySums(1)
    End If
    
    If myHowManyBlankLinesAfter > 0 Then
        myGrid.AddRow , , , , , , , myHowManyBlankLinesAfter
    End If
    
    myGrid.Redraw = True
    
End Function

Function CreateUnicodeFile(myPrinterType, myEAFDSSString, myInvoiceHeight, myDetailLines, myTopMargin, myLeftMargin)

    On Error GoTo ErrTrap
    
    Dim lngRow As Long
    Dim intProcessedDetailLines As Integer
    Dim intPageNo As Integer
    
    intPageNo = 0
    intProcessedDetailLines = 0
    
    Dim curTotals(1) As Currency
    
    Open strUnicodeFile For Output As #1
    InitReport myPrinterType, myEAFDSSString, myInvoiceHeight
    GoSub Headers
    
    With grdCommonPendingInvoices
        For lngRow = 1 To .RowCount
            Print #1, _
                Tab(1); .CellText(lngRow, "InvoiceIssueDate"); _
                Tab(12); .CellText(lngRow, "PersonDescription"); _
                Tab(63); .CellText(lngRow, "CodeDescription"); _
                Tab(110 - Len(.CellText(lngRow, "InvoiceNo"))); .CellText(lngRow, "InvoiceNo")
            intProcessedDetailLines = intProcessedDetailLines + 1
            If intProcessedDetailLines > myDetailLines Then
                If lngRow < .RowCount Then
                    Print #1, ""
                    Print #1, Space(11) & strMessages(24)
                    GoSub Headers
                    Print #1, Space(11) & strMessages(13)
                    Print #1, ""
                    intProcessedDetailLines = intProcessedDetailLines + 2
                End If
            End If
        Next lngRow
        DoEvents
    End With
    
    Print #1, ""
    Print #1, Space(11) & strMessages(25)
    
    Close #1
    
    CreateUnicodeFile = strUnicodeFile
    
    Exit Function
    
Headers:
    intPageNo = intPageNo + 1
    PrintHeadings 109, intPageNo, CustomUpperCase(lblTitle.Caption), CustomUpperCase(strCriteriaA), CustomUpperCase(strCriteriaB), myTopMargin
    PrintColumnHeadings 1, "ΗΜΕΡΟΜΗΝΙΑ", 12, "ΣΥΝΑΛΛΑΣΟΜΕΝΟΣ", 63, "ΠΑΡΑΣΤΑΤΙΚΟ", 108, "ΝΟ"
    Print #1, ""
    intProcessedDetailLines = 7
    
    Return
    
ErrTrap:
    Close #1
    CreateUnicodeFile = "Error"
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function TransformInvoices()

    If Not ValidateSelectedLines Then Exit Function
    
    If txtTriangularID.text = "1" Then
        txtInitialRefersTo.text = txtRefersTo.text
        txtRefersTo.text = "2"
    End If
    
    With CommonTransactions
        CustomizeGrid .grdCommonTransactions
        EnableFields .grdCommonTransactions
        EnableFields .mskTransDiscount, .mskTotalRestAmount, .mskExtraCharges, .mskTotalVAT
        InitializeFields IIf(.txtRefersTo.text = "2", .mskInvoiceIssueDate, ""), .mskTotalQty, .mskTotalPreDiscount, .mskDiscount, .mskTransDiscount, .mskTotalRestAmount, .mskExtraCharges, .mskTotalVAT, .mskTotalGross
        UpdateButtons CommonTransactions, 5, 0, 1, 0, 0, 1, 0
        .txtInvoiceDeliveryPointID.text = IIf(txtRefersTo.text = "1", "", "1")
        .txtInvoicePlates.Enabled = IIf(txtRefersTo.text = "1", False, True)
        .UpdateArrayWithInvoicesToTransform
        .UpdateRemarksFieldWithInvoices
        .UpdateGridWithItems
        .CalculateTotals True
        .txtRefersTo.text = txtRefersTo
        If txtTriangularID.text = "0" Then
            'Τιμολόγηση αγορών
            .UpdateHeaders
            .txtTable.text = txtTable.text
            .lblTitle.Caption = IIf(txtRefersTo = "1", "Τιμολόγηση αγορών", "Τιμολόγηση πωλήσεων")
            EnableFields .mskInvoiceIssueDate, .txtCodeShortDescription, .txtInvoiceNo, .txtInvoiceRemarks
            EnableFields .cmdIndex(2), .cmdIndex(3)
        Else
            'Τιμολόγηση τριγωνικών πωλήσεων
            .txtTable.text = "Customers"
            .lblTitle.Caption = "Τιμολόγηση τριγωνικών πωλήσεων"
            EnableFields .mskInvoiceIssueDate, .txtCodeShortDescription, .txtInvoiceNo, .txtInvoicePrintExtraRemarks, .txtInvoiceRemarks, .txtPersonDescription, .txtPaymentWayDescription
            EnableFields .cmdIndex(0), .cmdIndex(1), .cmdIndex(9), .cmdIndex(2), .cmdIndex(3), .cmdIndex(4), .cmdIndex(7), .cmdIndex(8)
            DisableFields .txtInvoicePlates
        End If
        .Tag = "True"
        .Show 1, Me
    End With

End Function

        
Private Function ValidateSelectedLines()

    'Local μεταβλητές
    Dim lngRow As Long
    Dim blnSelected As Boolean
    Dim strPersonDescription As String
    
    'Αρχικές τιμές
    ValidateSelectedLines = False
    blnSelected = False
    
    'Ελέγχω για επιλεγμένες γραμμές
    With grdCommonPendingInvoices
        For lngRow = 1 To .RowCount
            If .CellIcon(lngRow, "Selected") > 0 Then blnSelected = True: Exit For
        Next lngRow
    End With
    If Not blnSelected Then
        DisplayMessage 51, 4, 1, ""
        Exit Function
    End If

    'Ελέγχω για επιλεγμένες γραμμές του ίδιου συναλλασόμενου
    With grdCommonPendingInvoices
        For lngRow = 1 To .RowCount
            If .CellIcon(lngRow, "Selected") > 0 Then
                If strPersonDescription = "" Then
                    strPersonDescription = .CellText(lngRow, "PersonDescription")
                Else
                    If strPersonDescription <> .CellText(lngRow, "PersonDescription") Then
                        DisplayMessage 60, 4, 1, ""
                        Exit Function
                    End If
                End If
            End If
        Next lngRow
    End With

    'Τελικές τιμές
    ValidateSelectedLines = True

End Function


Private Function UpdateCriteriaLabels(myIssueFrom, myIssueTo, myInFrom, myInTo, myPerson)

    strCriteriaA = "Εκδοση από" & IIf(myIssueFrom <> "", " [ " & myIssueFrom & " ] ", " [ ΟΛΑ ] ") & "έως" & IIf(myIssueTo <> "", " [ " & myIssueTo & " ]", " [ ΟΛΑ ]")
    strCriteriaA = strCriteriaA & " Καταχώρηση από" & IIf(myInFrom <> "", " [ " & myInFrom & " ] ", " [ ΟΛΑ ] ") & "έως" & IIf(myInTo <> "", " [ " & myInTo & " ]", " [ ΟΛΑ ]")
    
    strCriteriaB = "Συναλλασόμενος" & IIf(myPerson <> "", " [ " & myPerson & " ] ", " [ ΟΛΟΙ ] ")
    
    lblCriteria.Caption = strCriteriaA & " " & strCriteriaB
    
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

Private Sub chkCriteriaItemAnalysis_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckForArrows (KeyCode)

End Sub


Private Sub chkCriteriaItemAnalysis_KeyPress(KeyAscii As Integer)

    ValidateInput (KeyAscii)

End Sub


Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
        Case 0
            If ValidateFields Then FindRecordsAndPopulateGrid
        Case 1
            SeekAndEditRecord _
                grdCommonPendingInvoices.CellText(grdCommonPendingInvoices.CurRow, "InvoiceTrnID"), _
                IIf(txtRefersTo.text = "1", "Αγορές", "Πωλήσεις"), _
                IIf(txtRefersTo.text = "1", "Καρτέλα προμηθευτή", "Καρτέλα πελάτη"), _
                txtTable.text, _
                txtRefersTo.text, _
                "", _
                ""
        Case 2
            TransformInvoices
        Case 3
            QuickTransformInvoices
        Case 4
            PrintRecords Me, "Print", False, "PrinterPrintsReportsID"
        Case 5
            PrintRecords Me, "CreatePDF", True, "PrinterPrintsReportsID"
        Case 6
            AbortProcedure False
        Case 7
            AbortProcedure True
    End Select
    
End Sub

Private Function ValidateFields()

    ValidateFields = False
    
    'Εκδοση
    If DisplayMessage(14, 4, 1, "", mskIssueFrom.text, mskIssueTo.text) Then mskIssueFrom.SetFocus: Exit Function
    
    'Καταχώρηση
    If DisplayMessage(14, 4, 1, "", mskInFrom.text, mskInTo.text) Then mskInFrom.SetFocus: Exit Function
    
    ValidateFields = True

End Function

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function
    
    If Not blnStatus Then
        If txtTriangularID.text = "1" Then
            txtRefersTo.text = txtInitialRefersTo.text
        End If
        ClearFields grdCommonPendingInvoices, lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
        frmCriteria(0).Visible = True
        mskIssueFrom.SetFocus
        UpdateButtons Me, 7, 1, 0, 0, 0, 0, 0, 0, 1
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
    Dim rstItems As Recordset

    'Local μεταβλητές
    Dim lngRow As Long
    Dim lngCol As Long
    
    'Αρχικές τιμές
    ReDim curGrandTotal(1)
    intIndex = 0
    lngRowCount = 0
    frmCriteria(0).Visible = False
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Πλέγμα
    With grdCommonPendingInvoices
        .Clear
        .Editable = False
        .Redraw = False
        .RowMode = False
    End With
    
    'SQL
    strSQL = "SELECT InvoiceIssueDate, " & txtTable.text & ".Description, " & txtTable.text & ".ID, " & "Codes.CodeDescription, Codes.CodeRefersTo, InvoiceNo, InvoiceTrnID, InvoiceID, InvoicePersonID, InvoiceRefersToID, InvoiceInDate, InvoiceDeliveryPointID, DeliveryPointDescription, InvoicePaymentWayID, PaymentWayDescription, InvoiceRemarks " _
    & "FROM ((((Invoices " _
    & "INNER JOIN " & txtTable.text & " ON Invoices.InvoicePersonID = " & txtTable.text & ".ID) " _
    & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID) " _
    & "INNER JOIN DeliveryPoints ON Invoices.InvoiceDeliveryPointID = DeliveryPoints.DeliveryPointID) " _
    & "INNER JOIN PaymentWays ON Invoices.InvoicePaymentWayID = PaymentWays.PaymentWayID) "

    'Αγορές ή Πωλήσεις
    strThisParameter = "intInvoiceRefersToID Integer"
    strThisQuery = "Invoices.InvoiceRefersToID = intInvoiceRefersToID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtRefersTo.text)
    
    'Τόπος παραλαβής
    If txtDeliveryPointID.text <> "" Then
        strThisParameter = "intDeliveryPointID Integer"
        strThisQuery = "Invoices.InvoiceDeliveryPointID = intDeliveryPointID "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtDeliveryPointID.text)
    End If
    
    'Εκδοση
    If IsDate(mskIssueFrom.text) Then
        strThisParameter = "datIssueFrom Date"
        strThisQuery = "Invoices.InvoiceIssueDate >= datIssueFrom"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskIssueFrom.text)
    End If
    If IsDate(mskIssueTo.text) Then
        strThisParameter = "datIssueTo Date"
        strThisQuery = "Invoices.InvoiceIssueDate <= datIssueTo"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskIssueTo.text)
    End If
        
    'Καταχώρηση
    If IsDate(mskInFrom.text) Then
        strThisParameter = "datInFrom Date"
        strThisQuery = "Invoices.InvoiceInDate >= datInFrom"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskInFrom.text)
    End If
    If IsDate(mskInTo.text) Then
        strThisParameter = "datInTo Date"
        strThisQuery = "Invoices.InvoiceInDate <= datInTo"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskInTo.text)
    End If
        
    'Συναλλασόμενος
    If txtPersonID.text <> "" Then
        strThisParameter = "intPerson Integer"
        strThisQuery = "Invoices.InvoicePersonID = intPerson"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtPersonID.text)
    End If
    
    'Μετασχηματίζεται σε ανώτερο επίπεδο
    strThisParameter = "lngCodeTransformID Long"
    strThisQuery = "Codes.CodeTransformID = lngCodeTransformID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = 1
    
    'Εκκρεμή παραστατικά - δεν ισχύει για τριγωνικές
    If txtTriangularID.text = "0" Then
        strThisParameter = "intInvoiceIsInvoiced Integer"
        strThisQuery = "Invoices.InvoiceIsInvoiced = intInvoiceIsInvoiced"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = 1
    End If

    'Ταξινόμηση
    strOrder = " ORDER BY InvoiceIssueDate, InvoiceNo"
    
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
    UpdateButtons Me, 7, 0, 0, 0, 0, 0, 0, 1, 0
    cmdButton(6).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        Do While Not .EOF
            grdCommonPendingInvoices.AddRow
            lngRow = grdCommonPendingInvoices.RowCount
            grdCommonPendingInvoices.CellValue(lngRow, "AA") = lngRowCount + 1
            grdCommonPendingInvoices.CellValue(lngRow, "InvoiceID") = !InvoiceID
            grdCommonPendingInvoices.CellValue(lngRow, "InvoiceIssueDate") = !InvoiceIssueDate
            grdCommonPendingInvoices.CellValue(lngRow, "InvoiceInDate") = !InvoiceInDate
            grdCommonPendingInvoices.CellValue(lngRow, "PersonID") = !ID
            grdCommonPendingInvoices.CellValue(lngRow, "PersonDescription") = !Description
            grdCommonPendingInvoices.CellValue(lngRow, "CodeDescription") = !CodeDescription
            grdCommonPendingInvoices.CellValue(lngRow, "InvoiceNo") = !InvoiceNo
            grdCommonPendingInvoices.CellValue(lngRow, "InvoiceTrnID") = !InvoiceTrnID
            grdCommonPendingInvoices.CellValue(lngRow, "DeliveryPointID") = !InvoiceDeliveryPointID
            grdCommonPendingInvoices.CellValue(lngRow, "DeliveryPointDescription") = !DeliveryPointDescription
            grdCommonPendingInvoices.CellValue(lngRow, "PaymentWayID") = !InvoicePaymentWayID
            grdCommonPendingInvoices.CellValue(lngRow, "PaymentWayDescription") = !PaymentWayDescription
            grdCommonPendingInvoices.CellValue(lngRow, "InvoiceRemarks") = !InvoiceRemarks
            If chkCriteriaItemAnalysis.Value = 1 Then GoSub FindItems
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
        ClearFields grdCommonPendingInvoices
    Else
        RefreshList = rstRecordset.RecordCount
        blnProcessing = False
    End If
    
    'Τελικές ενέργειες
    cmdButton(6).Caption = "Νέα αναζήτηση"
    frmProgress.Visible = False
    
    Exit Function
    
FindItems:
    strSQL = "SELECT Qty, UnitPrice, TotalNetPostDiscount, ItemDescription, ManufacturerDescription, ManufacturerIsShownID " _
        & "FROM ((InvoicesTrn " _
        & "LEFT JOIN Items ON InvoicesTrn.ItemID = Items.ItemID) " _
        & "LEFT JOIN Manufacturers ON Items.ItemManufacturerID = Manufacturers.ManufacturerID) " _
        & "WHERE InvoiceTrnID = " & rstRecordset!InvoiceTrnID
    strOrder = " ORDER BY ID"
    TempQuery.SQL = strSQL & strOrder
    Set rstItems = TempQuery.OpenRecordset()
    With rstItems
        Do While Not .EOF
            grdCommonPendingInvoices.AddRow
            lngRow = lngRow + 1
            grdCommonPendingInvoices.CellFont(lngRow, "PersonDescription").Name = "Input"
            grdCommonPendingInvoices.CellFont(lngRow, "PersonDescription").Size = "11"
            'grdCommonPendingInvoices.CellValue(lngRow, "Qty") = !Qty
            grdCommonPendingInvoices.CellValue(lngRow, "PersonDescription") = Trim(!ItemDescription) & IIf(!ManufacturerIsShownID = 1, " " & !ManufacturerDescription & " ", " ") & format(!Qty, "#,##0") & " x " & format(!TotalNetPostDiscount / !Qty, "#,##0.00") & " = " & format(!TotalNetPostDiscount, "#,##0.00")
            grdCommonPendingInvoices.CellTextFlags(lngRow, "PersonDescription") = igTextNoClip Or igTextLeft
            For lngCol = 1 To grdCommonPendingInvoices.colCount
                grdCommonPendingInvoices.CellForeColor(lngRow, lngCol) = vbCyan
            Next lngCol
            .MoveNext
        Loop
    End With
    
    Return
    
UpdateSQLString:
    intIndex = intIndex + 1
    strParameters = IIf(intIndex > 1, strParameters & ", ", strParameters)
    strParFields = IIf(intIndex > 1, strParFields & strLogic, strParFields)
    strParameters = strParameters & strThisParameter
    strParFields = strParFields & strThisQuery
    ReDim Preserve arrQuery(intIndex)
    
    Return
    
ErrTrap:
    If Err.Number = 6 Then Err.Description = Err.Description & " ID εγγραφής: " & rstRecordset!InvoiceID
    blnError = True
    ClearFields grdCommonPendingInvoices, frmProgress
    cmdButton(6).Caption = "Νέα αναζήτηση"
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function UpdateNextWindowTitle(myRefersToID)

    Select Case myRefersToID
        Case Is = 0
            UpdateNextWindowTitle = "Καρτέλα προμηθευτή"
        Case Is = 1
            UpdateNextWindowTitle = "Καρτέλα πελάτη"
        Case Is = 2
            UpdateNextWindowTitle = "Καρτέλα προμηθευτή"
        Case Is = 3
            UpdateNextWindowTitle = "Καρτέλα πελάτη"
    End Select

End Function




Private Sub cmdIndex_Click(Index As Integer)

    Dim strCategoryCriteria As String
    
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case Index
        Case 0
            'Προμηθευτής
            If txtPersonDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtPersonDescription.text, txtTable.text, "Description", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 3, 0, 1, 2, "ID", "Περιγραφή", "Α.Φ.Μ.", 0, 50, 15, 1, 0, 1)
            txtPersonID.text = tmpTableData.strCode
            txtPersonDescription.text = tmpTableData.strOneField
        Case 1
            'Τόπος παραλαβής
            If txtDeliveryPointDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtDeliveryPointDescription.text, "DeliveryPoints", "DeliveryPointDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtDeliveryPointID.text = tmpTableData.strCode
            txtDeliveryPointDescription.text = tmpTableData.strOneField
    End Select

End Sub

Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdCommonPendingInvoices, 44, GetSetting(strAppTitle, "Layout Strings", "grdCommonPendingInvoices" & txtRefersTo.text), _
            "05NCNAA,05NCNInvoiceID,05NCNInvoiceTrnID,10NCDXInvoiceIssueDate,10NCDXInvoiceInDate,10NCNPersonID,50NLNPersonDescription,40NLNCodeDescription,10NCNXInvoiceNo,10NCNDeliveryPointID,10LNNDeliveryPointDescription,10NCNPaymentWayID,10LNNPaymentWayDescription,05NCNOrder,05NCNInvoiceRemarks,03NCNSelected", _
            "A/A,InvoiceID,InvoiceTrnID,Ημερομηνία έκδοσης,Ημερομηνία καταχώρησης,ID Συναλλασόμενου,Συναλλασόμενος,Παραστατικό,Νο παραστατικού,ID Σημείου παράδοσης,Σημείο παράδοσης,ID Τρόπου πληρωμής,Τρόπος πληρωμής,Ταξινόμηση,Παρατηρήσεις,Ε"
        Me.Refresh
        frmCriteria(0).Visible = True
        mskIssueFrom.SetFocus
    End If
    
    'AddDummyLines grdCommonPendingInvoices, 5, 5, 5, 10, 12, 5, 50, 50, 6, 5, 40, 5, 40, 5, 40, 3
    
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
        Case vbKeyT And CtrlDown = 4 And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyT And CtrlDown = 8 And cmdButton(3).Enabled
            cmdButton_Click 3
        Case vbKeyP And CtrlDown = 4 And cmdButton(4).Enabled
            cmdButton_Click 4
        Case vbKeyP And CtrlDown = 8 And cmdButton(5).Enabled
            cmdButton_Click 5
        Case vbKeyEscape
            If cmdButton(6).Enabled Then cmdButton_Click 6: Exit Function
            If cmdButton(7).Enabled Then cmdButton_Click 7
        Case vbKeyF12
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    SetUpGrid lstIconList, grdCommonPendingInvoices
    PositionControls Me, True, grdCommonPendingInvoices
    ColorizeControls Me, True
    ClearFields mskIssueFrom, mskIssueTo, mskInFrom, mskInTo, txtPersonID, txtPersonDescription, txtDeliveryPointID, txtDeliveryPointDescription, lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals, chkCriteriaItemAnalysis
    UpdateButtons Me, 7, 1, 0, 0, 0, 0, 0, 0, 1

End Sub

Private Sub grdCommonPendingInvoices_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    If grdCommonPendingInvoices.RowCount = 0 Then Exit Sub
    
End Sub

Private Sub grdCommonPendingInvoices_ColHeaderMouseEnter(ByVal lCol As Long)

    grdCommonPendingInvoices.Header.Buttons = True

End Sub

Private Sub grdCommonPendingInvoices_ColHeaderMouseLeave(ByVal lCol As Long)

    grdCommonPendingInvoices.Header.Buttons = False
    
End Sub

Private Sub grdCommonPendingInvoices_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)

    cmdButton(1).Enabled = CheckToEnableButton(grdCommonPendingInvoices, lRow, "InvoiceID")

End Sub

Private Sub grdCommonPendingInvoices_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1
    
End Sub

Private Sub grdCommonPendingInvoices_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdCommonPendingInvoices_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    Dim lngRow As Long
    
    lngRow = grdCommonPendingInvoices.CurRow
    
    If KeyCode = vbKeyInsert Or KeyCode = vbKeyDelete Or KeyCode = vbKeySpace Then
        grdCommonPendingInvoices.CellIcon(grdCommonPendingInvoices.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdCommonPendingInvoices, KeyCode, grdCommonPendingInvoices.CurRow, "InvoiceID"))
        grdCommonPendingInvoices.CellValue(lngRow, "Order") = FindLastNumber(lngRow, grdCommonPendingInvoices.CellValue(lngRow, "Selected"))
        lblSelectedGridLines.Caption = CountSelected(grdCommonPendingInvoices)
    End If

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdCommonPendingInvoices" & txtRefersTo.text, grdCommonPendingInvoices.LayoutCol

End Sub

Private Sub txtDeliveryPointDescription_Change()

    If txtDeliveryPointDescription.text = "" Then ClearFields txtDeliveryPointID

End Sub

Private Sub txtDeliveryPointDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1

End Sub


Private Sub txtDeliveryPointDescription_Validate(Cancel As Boolean)

    If txtDeliveryPointID.text = "" And txtDeliveryPointDescription.text <> "" Then cmdIndex_Click 1: If txtDeliveryPointID.text = "" Then Cancel = True
    
End Sub

Private Sub txtPersonDescription_Change()

    If txtPersonDescription.text = "" Then ClearFields txtPersonID

End Sub

Private Sub txtPersonDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub

Private Sub txtPersonDescription_Validate(Cancel As Boolean)

    If txtPersonID.text = "" And txtPersonDescription.text <> "" Then cmdIndex_Click 0: If txtPersonID.text = "" Then Cancel = True

End Sub

Private Function FindLastNumber(lngRow, strNumber)

    With grdCommonPendingInvoices
        If .CellIcon(lngRow, "Selected") = 0 Then .CellValue(lngRow, "Order") = "": Exit Function
        For lngRow = 1 To .RowCount
            If .CellIcon(lngRow, "Selected") > 0 And .CellValue(lngRow, "Order") = "" Then intLastNo = intLastNo + 1
        Next lngRow
    End With

    FindLastNumber = intLastNo
    
End Function


