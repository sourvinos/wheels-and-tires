VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form ItemsInventory 
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
      Left            =   14550
      TabIndex        =   35
      Top             =   3600
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "ItemsInventory.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "ItemsInventory.frx":001C
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
         TabIndex        =   37
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
         Height          =   3840
         Index           =   1
         Left            =   150
         TabIndex        =   49
         Top             =   4200
         Visible         =   0   'False
         Width           =   9690
         Begin VB.Frame Frame1 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Height          =   540
            Left            =   2400
            TabIndex        =   64
            Top             =   3225
            Width           =   4890
            Begin GurhanButtonOCX.GurhanButton cmdButton 
               Height          =   390
               Index           =   10
               Left            =   2475
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   75
               Width           =   2190
               _ExtentX        =   3863
               _ExtentY        =   688
               Caption         =   "Ακυρο"
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
               BackColor       =   12648447
            End
            Begin GurhanButtonOCX.GurhanButton cmdButton 
               Height          =   390
               Index           =   9
               Left            =   225
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   75
               Width           =   2190
               _ExtentX        =   3863
               _ExtentY        =   688
               Caption         =   "Αποθήκευση"
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
               BackColor       =   12648447
            End
         End
         Begin UserControls.newDate mskDate 
            Height          =   465
            Left            =   2625
            TabIndex        =   5
            Top             =   825
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
         Begin UserControls.newText txtCodeShortDescription 
            Height          =   465
            Index           =   0
            Left            =   2625
            TabIndex        =   7
            Top             =   1875
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
            Index           =   2
            Left            =   4200
            TabIndex        =   58
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
            PicNormal       =   "ItemsInventory.frx":0038
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   4
            Left            =   4650
            TabIndex        =   59
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
            PicNormal       =   "ItemsInventory.frx":05D2
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newText txtCodeShortDescription 
            Height          =   465
            Index           =   1
            Left            =   2625
            TabIndex        =   8
            Top             =   2400
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
            Index           =   3
            Left            =   4200
            TabIndex        =   61
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
            PicNormal       =   "ItemsInventory.frx":0B6C
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   5
            Left            =   4650
            TabIndex        =   62
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
            PicNormal       =   "ItemsInventory.frx":1106
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newText txtInvoiceNo 
            Height          =   465
            Left            =   2625
            TabIndex        =   6
            Top             =   1350
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
            Index           =   1
            Left            =   5100
            TabIndex        =   63
            Top             =   2475
            Width           =   4215
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
            Index           =   0
            Left            =   5100
            TabIndex        =   60
            Top             =   1950
            Width           =   4215
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000000C0&
            Caption         =   "Παραστατικό πιστώσεων"
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
            TabIndex        =   57
            Top             =   2475
            Width           =   1740
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000000C0&
            Caption         =   "Παραστατικό χρεώσεων"
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
            TabIndex        =   55
            Top             =   1950
            Width           =   1740
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808000&
            Caption         =   "Δημιουργία εγγραφών απογραφής"
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
            Index           =   2
            Left            =   150
            TabIndex        =   54
            Top             =   75
            Width           =   2565
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
            Height          =   690
            Index           =   1
            Left            =   0
            TabIndex        =   53
            Top             =   3150
            Width           =   9705
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   5
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
            Index           =   4
            Left            =   9225
            Top             =   2175
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   1440
            Index           =   3
            Left            =   2175
            Top             =   1125
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label Label3 
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
            Left            =   5025
            TabIndex        =   52
            Top             =   75
            Width           =   4515
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000000C0&
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
            TabIndex        =   51
            Top             =   1425
            Width           =   1740
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H000000C0&
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
            Index           =   1
            Left            =   450
            TabIndex        =   50
            Top             =   900
            Width           =   1740
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
            Index           =   5
            Left            =   0
            TabIndex        =   56
            Top             =   0
            Width           =   9705
         End
      End
      Begin VB.Frame frmFrameForGridButtons 
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   75
         TabIndex        =   46
         Top             =   8250
         Width           =   8040
         Begin GurhanButtonOCX.GurhanButton cmdButton 
            Height          =   390
            Index           =   8
            Left            =   4050
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   75
            Width           =   3690
            _ExtentX        =   6509
            _ExtentY        =   688
            Caption         =   "Δημιουργία εγγραφών απογραφής"
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
            BackColor       =   12648447
         End
         Begin GurhanButtonOCX.GurhanButton cmdButton 
            Height          =   390
            Index           =   7
            Left            =   300
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   75
            Width           =   3690
            _ExtentX        =   6509
            _ExtentY        =   688
            Caption         =   "Μηδενισμός υπολοίπων"
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
            BackColor       =   12648447
         End
      End
      Begin VB.Frame frmButtonFrame 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   75
         TabIndex        =   28
         Top             =   8850
         Width           =   10365
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
            Caption         =   "Καρτέλα"
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
            Index           =   2
            Left            =   3075
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
            Index           =   5
            Left            =   7350
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
            Index           =   3
            Left            =   4500
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
            Index           =   4
            Left            =   5925
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            Caption         =   "Επεξεργασία κατάστασης"
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
         Height          =   3690
         Left            =   9900
         TabIndex        =   17
         Tag             =   "Hidden"
         Top             =   975
         Visible         =   0   'False
         Width           =   4515
         Begin VB.TextBox txtInvoiceCodeID 
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
            Index           =   1
            Left            =   3675
            TabIndex        =   70
            TabStop         =   0   'False
            Text            =   "5"
            Top             =   1575
            Width           =   780
         End
         Begin VB.TextBox Text5 
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
            TabIndex        =   69
            TabStop         =   0   'False
            Text            =   "Invoices.InvoiceCodeID(1)"
            Top             =   1575
            Width           =   3540
         End
         Begin VB.TextBox txtInvoiceCodeID 
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
            Index           =   0
            Left            =   3675
            TabIndex        =   68
            TabStop         =   0   'False
            Text            =   "4"
            Top             =   1200
            Width           =   780
         End
         Begin VB.TextBox Text2 
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
            Text            =   "Invoices.InvoiceCodeID(0)"
            Top             =   1200
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
            TabIndex        =   42
            TabStop         =   0   'False
            Text            =   "2"
            Top             =   450
            Width           =   780
         End
         Begin VB.TextBox Text4 
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
            TabIndex        =   41
            TabStop         =   0   'False
            Text            =   "Categories.CategoryID"
            Top             =   450
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
            Text            =   "Options.OptionID"
            Top             =   825
            Width           =   3540
         End
         Begin VB.TextBox txtOptionID 
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
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   "Table"
            Top             =   75
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
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   "1"
            Top             =   75
            Width           =   780
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   1950
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   2296
            Images          =   "ItemsInventory.frx":16A0
            Version         =   131072
            KeyCount        =   2
            Keys            =   "ÿ"
         End
      End
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   3315
         Index           =   0
         Left            =   9900
         TabIndex        =   9
         Top             =   4725
         Width           =   8865
         Begin VB.CheckBox chkCriteriaOnlyActiveItems 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Caption         =   "Μόνο τα ενεργά είδη"
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
            Left            =   1800
            TabIndex        =   4
            Top             =   2400
            Width           =   4065
         End
         Begin UserControls.newText txtOptionDescription 
            Height          =   465
            Left            =   1800
            TabIndex        =   2
            Top             =   1350
            Width           =   6165
            _ExtentX        =   10874
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
         Begin UserControls.newDate mskIssueTo 
            Height          =   465
            Left            =   1800
            TabIndex        =   3
            Top             =   1875
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
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   1
            Left            =   8025
            TabIndex        =   22
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
            PicNormal       =   "ItemsInventory.frx":1FB8
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newText txtCategoryShortDescription 
            Height          =   465
            Left            =   1800
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
            Left            =   2475
            TabIndex        =   38
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
            PicNormal       =   "ItemsInventory.frx":2552
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H000000C0&
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
            TabIndex        =   40
            Top             =   900
            Width           =   915
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
            Left            =   2925
            TabIndex        =   39
            Top             =   900
            Width           =   4365
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000000C0&
            Caption         =   "Εγγραφές"
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
            TabIndex        =   21
            Top             =   1425
            Width           =   915
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
            Left            =   4200
            TabIndex        =   18
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
            Left            =   1350
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
            Index           =   0
            Left            =   8400
            Top             =   1125
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
            TabIndex        =   16
            Top             =   2850
            Width           =   8880
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
            TabIndex        =   14
            Top             =   75
            Width           =   1665
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000000C0&
            Caption         =   "Εκδοση έως"
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
            TabIndex        =   10
            Top             =   1950
            Width           =   915
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
            TabIndex        =   15
            Top             =   0
            Width           =   8880
         End
      End
      Begin iGrid300_10Tec.iGrid grdItemsInventory 
         Height          =   6615
         Left            =   75
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1500
         Width           =   17415
         _ExtentX        =   30718
         _ExtentY        =   11668
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
      Begin VB.Label lblTotals 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Σύνολο κόστους"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   1
         Left            =   17700
         TabIndex        =   45
         Top             =   825
         Width           =   1215
      End
      Begin VB.Label lblTotals 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "-9.999.999,99"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   0
         Left            =   2550
         TabIndex        =   44
         Top             =   825
         Width           =   14940
      End
      Begin VB.Label lblSelectedGridTotals 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Σύνολα επιλεγμένων εγγραφών"
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
         Top             =   225
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
         TabIndex        =   20
         Top             =   525
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
         ForeColor       =   &H00FFFF00&
         Height          =   315
         Left            =   75
         TabIndex        =   19
         Top             =   1125
         Width           =   2565
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Υπόλοιπα ειδών"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   30
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   720
         Left            =   75
         TabIndex        =   13
         Top             =   75
         Width           =   3495
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
         TabIndex        =   12
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
Attribute VB_Name = "ItemsInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngRowCount As Long
Dim blnError As Boolean
Dim blnProcessing As Boolean
Dim blnEditingGrid As Boolean
Dim strCriteriaA As String
Dim strCriteriaB As String

Dim curCurrentQty() As Currency

Dim curLastCost As Currency
Dim datLastInvoiceIssueDate As Date
Dim strDefaultGridHeaderText(1) As String
Dim strUpdatedGridHeaderText(1) As String

Dim lngTrnID As Long
Private Function CalculateNewQtyBalanceAndCost(myGrid As iGrid, myRow As Long)

    Dim lngNewBalance As Long
    Dim curNewTotalCost As Currency
    
    lngNewBalance = Val(myGrid.CellValue(myRow, "CurrentQtyBalance")) + Val(myGrid.CellValue(myRow, "QtyPlus")) - Val(myGrid.CellValue(myRow, "QtyMinus"))
    myGrid.CellValue(myRow, "NewQtyBalance") = lngNewBalance
    
    curNewTotalCost = CCur(myGrid.CellValue(myRow, "LastBuyPrice")) * lngNewBalance
    myGrid.CellValue(myRow, "TotalCost") = curNewTotalCost
    
End Function

Private Function CalculateNewQtyTotalAndNewCostTotal(myGrid As iGrid, myFirstTime As Boolean)

    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngNewQtyBalance As Long
    Dim curNewTotalCost As Currency
    
    Dim intStep As Integer
    Dim lngDelay As Long
    Dim curCurrentTotalCost As Currency
    
    curCurrentTotalCost = IIf(lblTotals(0).Caption <> "", lblTotals(0).Caption, 0)
    
    For lngRow = 1 To myGrid.RowCount - 2
        lngNewQtyBalance = lngNewQtyBalance + Val(myGrid.CellValue(lngRow, "NewQtyBalance"))
        curNewTotalCost = curNewTotalCost + CCur(myGrid.CellValue(lngRow, "TotalCost"))
    Next lngRow
    
    myGrid.CellValue(myGrid.RowCount, "NewQtyBalance") = lngNewQtyBalance
    myGrid.CellValue(myGrid.RowCount, "TotalCost") = curNewTotalCost
    
    intStep = IIf(curNewTotalCost > curCurrentTotalCost, 10, -10)
    
    lblTotals(1).Caption = "Σύνολο κόστους"
    
    If Not myFirstTime Then
        For curCurrentTotalCost = curCurrentTotalCost To curNewTotalCost Step intStep
            lngDelay = 0
            lblTotals(0).Caption = format(curCurrentTotalCost, "#,##0.00")
            lblTotals(0).Refresh
            While lngDelay < 1000
                lngDelay = lngDelay + 1
                DoEvents
            Wend
        Next curCurrentTotalCost
    End If
    
    lblTotals(0).Caption = format(curNewTotalCost, "#,##0.00")

End Function

Private Function CreateInventoryRecords()

    On Error GoTo ErrTrap
    
    If Not ValidateFieldsForInventoryCreation Then Exit Function
    
    blnError = False
    
    BeginTrans
    
    SaveInvoice txtInvoiceCodeID(0).text 'Χρεώσεις
    SaveInvoicesTrn "QtyPlus"
    
    SaveInvoice txtInvoiceCodeID(1).text 'Πιστώσεις
    SaveInvoicesTrn "QtyMinus"
    
    If Not blnError Then
        CommitTrans
        frmCriteria(1).Visible = False
        ClearFields grdItemsInventory, mskDate, txtInvoiceNo, txtInvoiceCodeID(0), txtCodeShortDescription(0), lblCodeDescription(0), txtInvoiceCodeID(1), txtCodeShortDescription(1), lblCodeDescription(1)
        UpdateButtons Me, 10, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0
        If DisplayMessage(10, 1, 1, "", "") Then
            cmdButton_Click 0
        End If
    Else
        Rollback
    End If

    Exit Function
    
ErrTrap:
    Close #1
    CreateInventoryRecords = "Error"
    DisplayErrorMessage True, Err.Description

End Function

Private Function SaveInvoicesTrn(myColumn)

    Dim lngRow As Long
    
    If blnError Then Exit Function
    
    With grdItemsInventory
        For lngRow = 1 To .RowCount
            If .CellValue(lngRow, myColumn) <> "" Then
                If Not MainSaveRecord("CommonDB", "InvoicesTrn", True, strAppTitle, "InvoiceTrnID", lngTrnID, _
                    .CellValue(lngRow, "ItemID"), _
                    .CellValue(lngRow, myColumn), _
                    0, _
                    0, _
                    0, _
                    0, _
                    0, _
                    0, _
                    .CellValue(lngRow, "ItemVATPercent"), _
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


Private Function SaveInvoice(myDebitOrCredit)

    Dim lngRow As Long
    
    If blnError Then Exit Function
    
    lngTrnID = AddOneToTheLastRecord
    
    If Not MainSaveRecord("CommonDB", "Invoices", True, strAppTitle, "TrnID", _
        lngTrnID, _
        mskDate.text, Val(txtInvoiceNo.text), Val(myDebitOrCredit), 5, _
        0, 0, 0, 0, 0, 0, 0, 0, _
        lngTrnID, _
        "", _
        "", _
        6, _
        0, _
        0, _
        0, _
        Date, _
        Time, _
        "", _
        "", _
        "", _
        "", _
        1, _
        strCurrentUser) <> 0 Then
        blnError = True
    End If
    
End Function


Private Function ShowCreateInventoryFrame()

    frmCriteria(1).Visible = True
    UpdateButtons Me, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1
    mskDate.SetFocus

End Function

Private Function HideCreateInventoryFrame()

    frmCriteria(1).Visible = False
    UpdateButtons Me, 10, 0, 1, 0, 0, 0, 1, 0, 1, 1, 0, 0
    grdItemsInventory.SetFocus
    
End Function

Private Function FindRecordsAndPopulateGrid()

    If RefreshList > 0 Then
        UpdateRecordCount lblRecordCount, lngRowCount
        UpdateCriteriaLabels lblCategoryDescription, mskIssueTo.text, txtOptionDescription.text
        AddGridRowWithTotals grdItemsInventory, 0, "ItemDescription", strMessages(32), curGrandTotal(), 5, 2, 0, "CurrentQtyDebit", "CurrentQtyCredit", "CurrentQtyBalance", "LastBuyPrice", "NewQtyBalance", "TotalCost"
        CalculateNewQtyTotalAndNewCostTotal grdItemsInventory, True
        ColorizeCells grdItemsInventory, grdItemsInventory.RowCount, "CurrentQtyDebit", "CurrentQtyCredit", "CurrentQtyBalance", "LastBuyPrice", "NewQtyBalance", "TotalCost"
        EnableGrid grdItemsInventory, False
        HighlightRow grdItemsInventory, 1, "", True
        UpdateButtons Me, 10, 0, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0
    Else
        UpdateButtons Me, 10, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0
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
    
    Dim curTotals(5) As Currency
    
    Open strUnicodeFile For Output As #1
    InitReport myPrinterType, myEAFDSSString, myInvoiceHeight
    GoSub Headers
    
    With grdItemsInventory
        For lngRow = 1 To .RowCount
            Print #1, _
                Tab(7 - Len(.CellText(lngRow, "ItemID"))); .CellText(lngRow, "ItemID"); _
                Tab(8); Left(.CellText(lngRow, "ItemDescription"), 40); _
                Tab(49); Left(.CellText(lngRow, "ManufacturerDescription"), 25); _
                Tab(83 - Len(.CellText(lngRow, "PreviousQtyBalance"))); .CellText(lngRow, "PreviousQtyBalance"); _
                Tab(91 - Len(.CellText(lngRow, "QtyIn"))); .CellText(lngRow, "QtyIn"); _
                Tab(105 - Len(.CellText(lngRow, "ValueIn"))); .CellText(lngRow, "ValueIn"); _
                Tab(113 - Len(.CellText(lngRow, "QtyOut"))); .CellText(lngRow, "QtyOut"); _
                Tab(126 - Len(.CellText(lngRow, "ValueOut"))); .CellText(lngRow, "ValueOut"); _
                Tab(136 - Len(.CellText(lngRow, "QtyBalance"))); .CellText(lngRow, "QtyBalance")
            '///
            DoRunningTotal curTotals, .CellText(lngRow, "PreviousQtyBalance"), .CellText(lngRow, "QtyIn"), .CellText(lngRow, "ValueIn"), .CellText(lngRow, "QtyOut"), .CellText(lngRow, "ValueOut"), .CellText(lngRow, "QtyBalance")
            '///
            intProcessedDetailLines = intProcessedDetailLines + 1
            If intProcessedDetailLines > Val(myDetailLines) Then
                Print #1, ""
                AddTotalsToOutputFile Space(7) & strMessages(30), curTotals(), "083IY,091IY,105FY,113IY,126FY,136IY"
                GoSub Headers
                AddTotalsToOutputFile Space(7) & strMessages(31), curTotals(), "083IY,091IY,105FY,113IY,126FY,136IY"
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
    PrintHeadings 135, intPageNo, CustomUpperCase(lblTitle.Caption), CustomUpperCase(strCriteriaA), CustomUpperCase(strCriteriaB), myTopMargin
    PrintColumnHeadings 5, "ID", 8, "ΠΕΡΙΓΡΑΦΗ", 49, "ΚΑΤΑΣΚΕΥΑΣΤΗΣ", 75, "ΠΡΟΗΓ/ΝΟ", 85, "----- ΕΙΣΑΓΩΓΕΣ ----", 107, "----- ΕΞΑΓΩΓΕΣ ----", 128, "ΥΠΟΛΟΙΠΟ"
    PrintColumnHeadings 75, "ΥΠΟΛΟΙΠΟ", 85, "ΠΟΣ/ΤΑ", 101, "ΑΞΙΑ", 107, "ΠΟΣ/ΤΑ", 122, "ΑΞΙΑ", 129, "ΠΟΣ/ΤΑΣ"
    Print #1, ""
    intProcessedDetailLines = 8
    
    Return
    
ErrTrap:
    Close #1
    CreateUnicodeFile = "Error"
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function ShowLedger(myRow)

    With ItemsLedger
        .txtCategoryID.text = txtCategoryID.text
        .txtCategoryShortDescription.text = txtCategoryShortDescription.text
        .lblCategoryDescription.Caption = lblCategoryDescription.Caption
        .txtManufacturerID.text = grdItemsInventory.CellValue(myRow, "ManufacturerID")
        .txtManufacturerDescription.text = grdItemsInventory.CellValue(myRow, "ManufacturerDescription")
        .txtItemID.text = grdItemsInventory.CellText(grdItemsInventory.CurRow, "ItemID")
        .txtItemDescription.text = grdItemsInventory.CellText(grdItemsInventory.CurRow, "ItemDescription")
        .txtTable.text = txtTable.text
        .Tag = "True"
        DisableFields .txtCategoryShortDescription, .txtManufacturerDescription, .txtItemDescription, .cmdIndex(0), .cmdIndex(1), .cmdIndex(2)
        .Show 1, Me
    End With
    
End Function

Private Function UpdateCriteriaLabels(myCategory, myIssueTo, myCriteria)

    strCriteriaA = "Κατηγορία [ " & myCategory & " ] Κριτήρια " & "[ " & myCriteria & " ]"
    strCriteriaB = "Εκδοση έως " & IIf(myIssueTo <> "", "[ " & myIssueTo & " ]", "[ ΟΛΑ ]")
    
    lblCriteria.Caption = strCriteriaA & " " & strCriteriaB
    
End Function


Private Function UpdateWindowTitle(myRefersToID)

    Select Case myRefersToID
        Case Is = 3
            UpdateWindowTitle = "Καρτέλα προμηθευτή"
        Case Is = 4
            UpdateWindowTitle = "Καρτέλα πελάτη"
    End Select

End Function

Private Function ValidateFieldsForInventoryCreation()

    ValidateFieldsForInventoryCreation = False
    
    'Ημερομηνία
    If DisplayMessage(1, 4, 1, "", mskDate.text) Then mskDate.SetFocus: Exit Function
    
    'Νο παραστατικού
    If DisplayMessage(1, 4, 1, "", txtInvoiceNo.text) Then txtInvoiceNo.SetFocus: Exit Function
    
    'Τύπος παραστατικού χρεώσεων
    If DisplayMessage(1, 4, 1, "", txtInvoiceCodeID(0).text) Then txtCodeShortDescription(0).SetFocus: Exit Function
    
    'Τύπος παραστατικού πιστώσεων
    If DisplayMessage(1, 4, 1, "", txtInvoiceCodeID(1).text) Then txtCodeShortDescription(1).SetFocus: Exit Function
    
    ValidateFieldsForInventoryCreation = True

End Function

Private Function ZeroQtyAndCost()

    Dim lngRow As Long
    
    UpdateButtons Me, 10, 0, 1, 0, 0, 0, 1, 0, 0, 1, 0, 0
    
    With grdItemsInventory
        For lngRow = 1 To .RowCount - 2
            'Αν το υπόλοιπο είναι < 0 η διαφορά πάει στη χρέωση
            If .CellValue(lngRow, "NewQtyBalance") < 0 Then
                .CellValue(lngRow, "QtyPlus") = Abs(.CellValue(lngRow, "NewQtyBalance"))
                CalculateNewQtyBalanceAndCost grdItemsInventory, lngRow
            End If
            'Αν το υπόλοιπο είναι > 0 η διαφορά πάει στην πίστωση
            If .CellValue(lngRow, "NewQtyBalance") > 0 Then
                .CellValue(lngRow, "QtyMinus") = .CellValue(lngRow, "NewQtyBalance")
                CalculateNewQtyBalanceAndCost grdItemsInventory, lngRow
            End If
            ColorizeCells grdItemsInventory, lngRow, "CurrentQtyDebit", "CurrentQtyCredit", "CurrentQtyBalance", "LastBuyPrice", "NewQtyBalance", "TotalCost"
        Next lngRow
        
        .SetFocus
    
    End With

    CalculateNewQtyTotalAndNewCostTotal grdItemsInventory, True
    ColorizeCells grdItemsInventory, grdItemsInventory.RowCount, "NewQtyBalance", "TotalCost"
    
End Function

Private Sub chkCriteriaOnlyActiveItems_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckForArrows (KeyCode)

End Sub


Private Sub chkCriteriaOnlyActiveItems_KeyPress(KeyAscii As Integer)

    ValidateInput (KeyAscii)

End Sub


Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
        Case 0
            If ValidateFields Then FindRecordsAndPopulateGrid
        Case 1
            ShowLedger grdItemsInventory.CurRow
        Case 2
            PrintRecords Me, "Print", False, "PrinterPrintsReportsID"
        Case 3
            PrintRecords Me, "CreatePDF", True, "PrinterPrintsReportsID"
        Case 4
            EditGrid
        Case 5
            AbortProcedure False
        Case 6
            AbortProcedure True
        Case 7
            ZeroQtyAndCost
        Case 8
            ShowCreateInventoryFrame
        Case 9
            CreateInventoryRecords
        Case 10
            AbortProcedure False
    End Select
    
End Sub

Private Function EditGrid()

    blnEditingGrid = True
    UpdateButtons Me, 10, 0, 1, 0, 0, 0, 1, 0, 1, 1, 0, 0
    cmdButton(5).Caption = "Ακυρο"
    EnableGrid grdItemsInventory, True, grdItemsInventory.CurRow, 11

End Function


Private Function ValidateFields()

    ValidateFields = False
    
    'Κατηγορία
    If DisplayMessage(1, 4, 1, "", txtCategoryID.text) Then txtCategoryShortDescription.SetFocus: Exit Function
    
    'Εγγραφές
    If DisplayMessage(1, 4, 1, "", txtOptionID.text) Then txtOptionDescription.SetFocus: Exit Function
    
    'Εως
    If DisplayMessage(1, 4, 1, "", mskIssueTo.text) Then mskIssueTo.SetFocus: Exit Function
    
    ValidateFields = True

End Function


Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function
    
    If frmCriteria(1).Visible Then
        If MyMsgBox(3, strAppTitle, strMessages(3), 2) Then
            HideCreateInventoryFrame
            Exit Function
        End If
        Exit Function
    End If
    
    If Not blnStatus Then
        If Not blnEditingGrid Then
            ClearFields grdItemsInventory, lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals, lblTotals(0), lblTotals(1)
            frmCriteria(0).Visible = True
            txtCategoryShortDescription.SetFocus
            UpdateButtons Me, 10, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0
        Else
            If MyMsgBox(3, strAppTitle, strMessages(3), 2) Then
                EnableGrid grdItemsInventory, False, grdItemsInventory.CurRow
                UpdateButtons Me, 10, 0, CheckToEnableButton(grdItemsInventory, grdItemsInventory.CurRow, "ItemID"), 1, 1, 1, 1, 0, 1, 1, 0, 0
                ClearNewInventoryQtyAndCost
                blnEditingGrid = False
            End If
        End If
    End If
    
    If blnStatus Then
        Unload Me
    End If

End Function

Private Function ClearNewInventoryQtyAndCost()

    Dim lngRow As Long
    
    With grdItemsInventory
        For lngRow = 1 To .RowCount - 2
            .CellValue(lngRow, "QtyPlus") = ""
            .CellValue(lngRow, "QtyMinus") = ""
            CalculateNewQtyBalanceAndCost grdItemsInventory, lngRow
            ColorizeCells grdItemsInventory, lngRow, "NewQtyBalance", "TotalCost"
        Next lngRow
    End With
    
    CalculateNewQtyTotalAndNewCostTotal grdItemsInventory, True
    ColorizeCells grdItemsInventory, grdItemsInventory.RowCount, "NewQtyBalance", "TotalCost"

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

    'Local μεταβλητές
    Dim intLoop As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngItemID As Long
    Dim strItemDescription As String
    Dim lngManufacturerID As Long
    Dim strManufacturerDescription As String
    Dim curItemVATPercent As Currency
    
    'Αρχικές τιμές
    ReDim curCurrentQty(2)
    ReDim curGrandTotal(5)
    
    intIndex = 0
    lngRowCount = 0
    frmCriteria(0).Visible = False
    blnEditingGrid = False
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Πλέγμα
    With grdItemsInventory
        .Clear
        .Editable = False
        .Redraw = False
        .RowMode = False
    End With
    
    'Αγορές, πωλήσεις, κινήσεις πελατών και προμηθευτών
    strSQL = "SELECT InvoiceIssueDate, InvoicesTrn.ItemID, Qty, TotalNetPostDiscount, InvoicesTrn.InvoiceTrnID, ItemDescription, ManufacturerID, ManufacturerDescription, CodeInventoryQty, CodeInventoryValue, InvoiceRefersToID, ItemVATPercent " _
    & "FROM (((InvoicesTrn " _
    & "INNER JOIN Invoices ON InvoicesTrn.InvoiceTrnID = Invoices.InvoiceTrnID) " _
    & "INNER JOIN Items ON InvoicesTrn.ItemID = Items.ItemID) " _
    & "INNER JOIN Manufacturers ON Items.ItemManufacturerID = Manufacturers.ManufacturerID) " _
    & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID "
    
    'Κατηγορία
    If txtCategoryID.text <> "" Then
        strThisParameter = "lngCategoryID Long"
        strThisQuery = "Items.ItemCategoryID = lngCategoryID"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtCategoryID.text)
    End If
    
    'Είδος
    'strThisParameter = "lngItemID Long"
    'strThisQuery = "Items.ItemID = lngItemID"
    'strLogic = " AND "
    'GoSub UpdateSQLString
    'arrQuery(intIndex) = "302"
    
    'Εως
    If IsDate(mskIssueTo.text) Then
        strThisParameter = "datIssueTo Date"
        strThisQuery = "Invoices.InvoiceIssueDate <= datIssueTo"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskIssueTo.text)
    End If
    
    'Μόνο τα ενεργά είδη
    If chkCriteriaOnlyActiveItems.Value = 1 Then
        strThisParameter = "strActiveItems String"
        strThisQuery = "Items.ItemActive = strActiveItems"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Trim(Str(chkCriteriaOnlyActiveItems.Value))
    End If
        
    'Ταξινόμηση
    strOrder = " ORDER BY InvoicesTrn.ItemID, InvoiceIssueDate"
    
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
    UpdateButtons Me, 10, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0
    cmdButton(5).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    '1η εγγραφή
    GoSub UpdateAreas
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        Do While Not .EOF
            If !ItemID = lngItemID Then
                CalculatePeriod lngItemID, rstRecordset
                If Not blnProcessing Then Exit Do
            Else
                If txtOptionID.text = "1" Or (txtOptionID.text = "2" And curCurrentQty(2) <> 0) Then
                    GoSub AddLine
                    ColorizeCells grdItemsInventory, lngRow, "CurrentQtyBalance", "LastBuyPrice", "NewQtyBalance", "TotalCost"
                    CalculateGrandTotals curCurrentQty(0), curCurrentQty(1), curCurrentQty(2), curCurrentQty(2) * curLastCost, curCurrentQty(2), curCurrentQty(2) * curLastCost
                End If
                ClearVariables curCurrentQty(0), curCurrentQty(1), curCurrentQty(2), curLastCost
                GoSub UpdateAreas
            End If
        Loop
        If blnProcessing Then
            If txtOptionID.text = "1" Or (txtOptionID.text = "2" And curCurrentQty(2) <> 0) Then
                GoSub AddLine
                ColorizeCells grdItemsInventory, lngRow, "CurrentQtyBalance", "LastBuyPrice", "NewQtyBalance", "TotalCost"
                CalculateGrandTotals curCurrentQty(0), curCurrentQty(1), curCurrentQty(2), curCurrentQty(2) * curLastCost, curCurrentQty(2), curCurrentQty(2) * curLastCost
            End If
        End If
    End With
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        RefreshList = 0
        ClearFields grdItemsInventory
    Else
        grdItemsInventory.Sort Array("ManufacturerDescription", "ItemDescription")
        RefreshList = lngRowCount
        blnProcessing = False
    End If
    
    'Τελικές ενέργειες
    cmdButton(5).Caption = "Νέα αναζήτηση"
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
    
AddLine:
    With grdItemsInventory
        .AddRow
        lngRow = .RowCount
        .CellValue(.RowCount, "ItemID") = lngItemID
        .CellValue(.RowCount, "ItemDescription") = strItemDescription
        .CellValue(.RowCount, "ItemVATPercent") = curItemVATPercent
        .CellValue(.RowCount, "ManufacturerID") = lngManufacturerID
        .CellValue(.RowCount, "ManufacturerDescription") = strManufacturerDescription
        .CellValue(.RowCount, "CurrentQtyDebit") = curCurrentQty(0)
        .CellValue(.RowCount, "CurrentQtyCredit") = curCurrentQty(1)
        .CellValue(.RowCount, "CurrentQtyBalance") = curCurrentQty(2)
        .CellValue(.RowCount, "LastBuyPrice") = curLastCost
        .CellValue(.RowCount, "NewQtyBalance") = curCurrentQty(2)
        .CellValue(.RowCount, "TotalCost") = curCurrentQty(2) * curLastCost
        lngRowCount = lngRowCount + 1
    End With
    
    Return

UpdateAreas:
    lngItemID = rstRecordset!ItemID
    strItemDescription = rstRecordset!ItemDescription
    curItemVATPercent = rstRecordset!ItemVATPercent
    lngManufacturerID = rstRecordset!ManufacturerID
    strManufacturerDescription = rstRecordset!ManufacturerDescription
    datLastInvoiceIssueDate = rstRecordset!InvoiceIssueDate
    
    Return

ErrTrap:
    blnError = True
    ClearFields grdItemsInventory, frmProgress
    cmdButton(5).Caption = "Νέα αναζήτηση"
    DisplayErrorMessage True, Err.Description
        
End Function

Private Function CalculatePeriod(myID, myRecordset As Recordset)

    With myRecordset
        Do While !InvoiceIssueDate <= CDate(mskIssueTo.text) And myID = !ItemID
            FillArray curCurrentQty, _
                CalculateDebitCreditAndBalance("Debit", "Items", !Qty, "", "", !CodeInventoryQty, "", ""), _
                CalculateDebitCreditAndBalance("Credit", "Items", !Qty, "", "", !CodeInventoryQty, "", "")
            UpdateLastCost myRecordset

            UpdateProgressBar Me
            .MoveNext
            DoEvents
            If .EOF Then
                Exit Do
            Else
                If Not blnProcessing Then Exit Function
                If !InvoiceIssueDate > CDate(mskIssueTo.text) Or !ItemID <> myID Then
                    Exit Do
                End If
            End If
        Loop
        curCurrentQty(2) = curCurrentQty(0) - curCurrentQty(1)
        CalculatePeriod = curCurrentQty()
    End With

End Function


Private Function UpdateLastCost(myRecordset As Recordset)

    With myRecordset
        If !InvoiceRefersToID = 0 And !CodeInventoryValue = "+" Then
            If !InvoiceIssueDate > datLastInvoiceIssueDate Then
                If !Qty > 0 Then
                    curLastCost = !TotalNetPostDiscount / !Qty
                End If
                datLastInvoiceIssueDate = !InvoiceIssueDate
            End If
        End If
    End With
    
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
            'Εγγραφές
            If txtOptionDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtOptionDescription.text, "Options", "OptionDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtOptionID.text = tmpTableData.strCode
            txtOptionDescription.text = tmpTableData.strOneField
        Case 2, 3
            'Παραστατικό χρεώσεων - πιστώσεων
            If txtCodeShortDescription(Index - 2).text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtCodeShortDescription(Index - 2).text, "Codes", "CodeShortDescription", "String", "5", 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 3, 0, 1, 2, "ID", "Συντ.", "Περιγραφή", 0, 4, 40, 1, 1, 0)
            txtInvoiceCodeID(Index - 2).text = tmpTableData.strCode
            txtCodeShortDescription(Index - 2).text = tmpTableData.strOneField
            lblCodeDescription(Index - 2).Caption = tmpTableData.strTwoField
        Case 4, 5
            'Παραστατικό χρεώσεων - πιστώσεων
            With UtilsCodes
                .Tag = "True"
                .txtRefersTo.text = "5"
                .Show 1, Me
            End With
    End Select

End Sub

Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        strDefaultGridHeaderText(0) = "Υπόλοιπο"
        strDefaultGridHeaderText(1) = "Τ. Τ. Α."
        strUpdatedGridHeaderText(0) = "Υπόλοιπο στις " & Chr(13)
        strUpdatedGridHeaderText(1) = "Τ.Τ.Α έως " & Chr(13)
        AddColumnsToGrid grdItemsInventory, 44, GetSetting(strAppTitle, "Layout Strings", "grdItemsInventory"), _
            "06NCNItemID,06NCNCategoryID, 40NLNCategoryDescription,10NCNItemVATPercent,50NLNItemDescription,06NCNManufacturerID,40NLNManufacturerDescription,10NRICurrentQtyDebit,10NRICurrentQtyCredit,10NRICurrentQtyBalance,10NRFLastBuyPrice,05NRIXQtyPlus,05NRIXQtyMinus,10NRIXNewQtyBalance,10NRFTotalCost,03NCNSelected", _
            "ID,ID Κατηγορίας,Κατηγορία,Φ.Π.Α.,Περιγραφή,ID Κατασκευαστή,Κατασκευαστής,Εισαγωγές,Εξαγωγές," & strDefaultGridHeaderText(0) & "," & strDefaultGridHeaderText(1) & ",Ποσότητα" & Chr(13) & "( + ),Ποσότητα " & Chr(13) & "( - ),Νέο υπόλοιπο,Κόστος,Ε"
        Me.Refresh
        frmCriteria(0).Visible = True
        frmCriteria(1).Visible = False
        txtCategoryShortDescription.SetFocus
    End If
    
    'AddDummyLines grdItemsInventory, 6, 6, 4, 40, 50, 6, 40, 11, 11, 10, 10, 11, 9, 10, 11, 3
    
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
        Case vbKeyL And CtrlDown = 4 And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyP And CtrlDown = 4 And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyP And CtrlDown = 8 And cmdButton(3).Enabled
            cmdButton_Click 3
        Case vbKeyEscape
            If cmdButton(5).Enabled Then cmdButton_Click 5: Exit Function
            If cmdButton(6).Enabled Then cmdButton_Click 6: Exit Function
            If cmdButton(10).Enabled Then cmdButton_Click 10
        Case vbKeyZ And CtrlDown = 4 And cmdButton(7).Enabled
            cmdButton_Click 7
        Case vbKeyN And CtrlDown = 4 And cmdButton(8).Enabled
            cmdButton_Click 8
        Case vbKeyF10 And cmdButton(9).Enabled, vbKeyS And CtrlDown = 4 And cmdButton(9).Enabled
            cmdButton_Click 9
        Case vbKeyF12
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    SetUpGrid lstIconList, grdItemsInventory
    PositionControls Me, True, grdItemsInventory
    ColorizeControls Me, True
    ClearFields txtCategoryID, txtCategoryShortDescription, lblCategoryDescription, txtOptionID, txtOptionDescription, mskIssueTo, chkCriteriaOnlyActiveItems, lblRecordCount, lblCriteria, lblTotals(0), lblTotals(1), lblSelectedGridLines, lblSelectedGridTotals, mskDate, txtInvoiceNo, txtCodeShortDescription(0), lblCodeDescription(0), txtInvoiceCodeID(0), txtCodeShortDescription(1), lblCodeDescription(1), txtInvoiceCodeID(1)
    UpdateButtons Me, 10, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0

End Sub

Private Sub grdItemsInventory_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)

    CalculateNewQtyBalanceAndCost grdItemsInventory, lRow
    CalculateNewQtyTotalAndNewCostTotal grdItemsInventory, False
    ColorizeCells grdItemsInventory, grdItemsInventory.CurRow, "NewQtyBalance", "TotalCost"
    ColorizeCells grdItemsInventory, grdItemsInventory.RowCount, "NewQtyBalance", "TotalCost"
    
End Sub


Private Sub grdItemsInventory_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    If grdItemsInventory.RowCount = 0 Then Exit Sub

    grdItemsInventory.RemoveRow (grdItemsInventory.RowCount): grdItemsInventory.RemoveRow (grdItemsInventory.RowCount)

End Sub

Private Sub grdItemsInventory_ColHeaderMouseEnter(ByVal lCol As Long)

    grdItemsInventory.Header.Buttons = True

End Sub

Private Sub grdItemsInventory_ColHeaderMouseLeave(ByVal lCol As Long)

    grdItemsInventory.Header.Buttons = False
    
End Sub

Private Sub grdItemsInventory_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)

    cmdButton(1).Enabled = CheckToEnableButton(grdItemsInventory, lRow, "ItemID")

End Sub

Private Sub grdItemsInventory_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1

End Sub


Private Sub grdItemsInventory_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdItemsInventory_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeyInsert Or KeyCode = vbKeyDelete Or KeyCode = vbKeySpace Then
        grdItemsInventory.CellIcon(grdItemsInventory.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdItemsInventory, KeyCode, grdItemsInventory.CurRow, "ItemID"))
        lblSelectedGridLines.Caption = CountSelected(grdItemsInventory)
        lblSelectedGridTotals.Caption = SumSelectedGridRows(grdItemsInventory, False, "NewQtyBalance", "TotalCost")
    End If

End Sub

Private Sub grdItemsInventory_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid300_10Tec.ETextEditFlags)

    If lCol = 11 Or lCol = 12 Then
        If CheckForAcceptableKey(iKeyAscii) Then
            CaptureNumbers grdItemsInventory.TextEditText, lRow, lCol, iKeyAscii, True
        Else
            bCancel = True
        End If
    Else
        bCancel = True
    End If

End Sub

Private Sub grdItemsInventory_TextEditKeyPress(ByVal lRow As Long, ByVal lCol As Long, KeyAscii As Integer)

    If lCol = 11 Or lCol = 12 Then
        If CheckForAcceptableKey(KeyAscii) Then
            CaptureNumbers grdItemsInventory.TextEditText, lRow, lCol, KeyAscii, True
        Else
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdItemsInventory", grdItemsInventory.LayoutCol

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


Private Sub txtCodeShortDescription_Change(Index As Integer)

    If txtCodeShortDescription(Index).text = "" Then ClearFields txtInvoiceCodeID(Index), lblCodeDescription(Index)

End Sub

Private Sub txtCodeShortDescription_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If Index = 0 And KeyCode = vbKeyF2 Then cmdIndex_Click 2
    If Index = 1 And KeyCode = vbKeyF2 Then cmdIndex_Click 3
    
    If Index = 0 And KeyCode = vbKeyF5 Then cmdIndex_Click 4
    If Index = 1 And KeyCode = vbKeyF5 Then cmdIndex_Click 5
    
End Sub


Private Sub txtCodeShortDescription_Validate(Index As Integer, Cancel As Boolean)

    If txtInvoiceCodeID(Index).text = "" And txtCodeShortDescription(Index).text <> "" Then cmdIndex_Click Index + 2: If txtInvoiceCodeID(Index).text = "" Then Cancel = True

End Sub


Private Sub txtOptionDescription_Change()

    If txtOptionDescription.text = "" Then ClearFields txtOptionID
    
End Sub

Private Sub txtOptionDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1

End Sub

Private Sub txtOptionDescription_Validate(Cancel As Boolean)

    If txtOptionID.text = "" And txtOptionDescription.text <> "" Then cmdIndex_Click 1: If txtOptionID.text = "" Then Cancel = True

End Sub

