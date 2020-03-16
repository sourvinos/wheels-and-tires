VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form SalesIncomingVehicles 
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
      TabIndex        =   10
      Top             =   7650
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "SalesIncomingVehicles.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "SalesIncomingVehicles.frx":001C
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
         TabIndex        =   12
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
         TabIndex        =   26
         Top             =   8850
         Width           =   8940
         Begin GurhanButtonOCX.GurhanButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   27
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
            TabIndex        =   28
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
            Index           =   5
            Left            =   7350
            TabIndex        =   29
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
            TabIndex        =   30
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
            Index           =   4
            Left            =   5925
            TabIndex        =   31
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
            TabIndex        =   32
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
      End
      Begin VB.Frame frmInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   9450
         TabIndex        =   16
         Tag             =   "Hidden"
         Top             =   5700
         Visible         =   0   'False
         Width           =   4515
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
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "3"
            Top             =   825
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
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   "RefersTo"
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
            TabIndex        =   34
            TabStop         =   0   'False
            Text            =   "2"
            Top             =   450
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
            TabIndex        =   33
            TabStop         =   0   'False
            Text            =   "Table"
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
            TabIndex        =   24
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
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   "1"
            Top             =   75
            Width           =   780
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   1200
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   2296
            Images          =   "SalesIncomingVehicles.frx":0038
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
         TabIndex        =   5
         Top             =   5625
         Width           =   9240
         Begin UserControls.newText txtPersonDescription 
            Height          =   465
            Left            =   2175
            TabIndex        =   3
            Top             =   1350
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
            PicNormal       =   "SalesIncomingVehicles.frx":0950
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newText txtPlates 
            Height          =   465
            Left            =   2175
            TabIndex        =   4
            Top             =   1875
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   20
            Text            =   "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ"
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
            BackColor       =   &H000000C0&
            Caption         =   "Αρ. κυκλοφορίας"
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
            Top             =   1950
            Width           =   1290
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000000C0&
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
            Index           =   0
            Left            =   450
            TabIndex        =   20
            Top             =   1425
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
            TabIndex        =   17
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
            TabIndex        =   15
            Top             =   2625
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
            TabIndex        =   13
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
            TabIndex        =   6
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
            TabIndex        =   14
            Top             =   0
            Width           =   9240
         End
      End
      Begin iGrid300_10Tec.iGrid grdSalesIncomingVehicles 
         Height          =   7290
         Left            =   75
         TabIndex        =   7
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
         TabIndex        =   25
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
         TabIndex        =   19
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
         ForeColor       =   &H00FFFF00&
         Height          =   315
         Left            =   75
         TabIndex        =   18
         Top             =   1125
         Width           =   2565
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ημερολόγιο εισερχομένων οχημάτων"
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
         Left            =   75
         TabIndex        =   9
         Top             =   75
         Width           =   8685
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
         TabIndex        =   8
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
Attribute VB_Name = "SalesIncomingVehicles"
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

Private Function SeekAndEditRecord(myInvoiceTrnID, myWindowTitle, myNextWindowTitle, myTable, myRefersTo)
    
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
        UpdateCriteriaLabels mskIssueFrom.text, mskIssueTo.text, txtPersonDescription.text, txtPlates.text
        AddGridRowWithTotals grdSalesIncomingVehicles, 0, "PersonDescription", False, strMessages(32), True, curGrandTotal(), 1, 2, 0, "InvoiceGrossAmount"
        EnableGrid grdSalesIncomingVehicles, False
        HighlightRow grdSalesIncomingVehicles, 1, "", True
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
    
    With grdSalesIncomingVehicles
        For lngRow = 1 To .RowCount - 2
            Print #1, _
                Tab(1); .CellText(lngRow, "InvoiceIssueDate"); _
                Tab(12); .CellText(lngRow, "PersonDescription"); _
                Tab(53); Left(.CellText(lngRow, "Address"), 30); _
                Tab(84); .CellText(lngRow, "Plates"); _
                Tab(105); .CellText(lngRow, "CodeShortDescription"); _
                Tab(110); .CellText(lngRow, "InvoiceNo"); _
                Tab(117); .CellText(lngRow, "InvoiceInTime"), _
                Tab(136 - Len(.CellText(lngRow, "InvoiceGrossAmount"))); .CellText(lngRow, "InvoiceGrossAmount")
            DoRunningTotal curTotals, .CellText(lngRow, "InvoiceGrossAmount")
            intProcessedDetailLines = intProcessedDetailLines + 1
            If intProcessedDetailLines > myDetailLines Then
                If lngRow < .RowCount - 2 Then
                    Print #1, ""
                    AddTotalsToOutputFile Space(11) & strMessages(30), curTotals(), "136FY"
                    GoSub Headers
                    AddTotalsToOutputFile Space(11) & strMessages(31), curTotals(), "136FY"
                    Print #1, '"
                    intProcessedDetailLines = intProcessedDetailLines + 2
                End If
            End If
        Next lngRow
        DoEvents
    End With
    
    Print #1, ""
    AddTotalsToOutputFile Space(11) & strMessages(32), curTotals(), "136FY"
    
    Close #1
    
    CreateUnicodeFile = strUnicodeFile
    
    Exit Function
    
Headers:
    intPageNo = intPageNo + 1
    PrintHeadings 135, intPageNo, CustomUpperCase(lblTitle.Caption), CustomUpperCase(strCriteriaA), CustomUpperCase(strCriteriaB), myTopMargin
    PrintColumnHeadings 1, "ΗΜΕΡΟΜΗΝΙΑ", 12, "ΣΥΝΑΛΛΑΣΟΜΕΝΟΣ", 53, "ΔΙΕΥΘΥΝΣΗ", 84, "ΑΡ. ΚΥΚΛΟΦΟΡΙΑΣ", 105, "ΠΑΡΑΣΤΑΤΙΚΟ", 117, "ΩΡΑ", 128, "ΣΥΝΟΛΙΚΗ"
    PrintColumnHeadings 117, "ΕΞΟΔΟΥ", 132, "ΑΞΙΑ"
    Print #1, ""
    intProcessedDetailLines = 8
    
    Return
    
ErrTrap:
    Close #1
    CreateUnicodeFile = "Error"
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function UpdateCriteriaLabels(myIssueFrom, myIssueTo, myPerson, myPlates)

    strCriteriaA = "Εκδοση από" & IIf(myIssueFrom <> "", " [ " & myIssueFrom & " ] ", " [ ΟΛΑ ] ") & "έως" & IIf(myIssueTo <> "", " [ " & myIssueTo & " ]", " [ ΟΛΑ ]")
    strCriteriaB = "Επωνυμία" & IIf(myPerson <> "", " [ " & myPerson & " ] ", " [ ΟΛΟΙ ] ")
    
    If myPlates = "" Then
        strCriteriaB = strCriteriaB & "Αρ. κυκλοφορίας [ ΟΛΟΙ ]"
    Else
        If Left(myPlates, 1) <> "*" Then strCriteriaB = strCriteriaB & "Αρ. κυκλοφορίας αρχίζει από [ " & UCase(myPlates) & " ]"
        If Left(myPlates, 1) = "*" Then strCriteriaB = strCriteriaB & "Αρ. κυκλοφορίας περιέχει το [ " & UCase(Right(myPlates, Len(myPlates) - 1)) & " ]"
    End If
    
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

Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
        Case 0
            If ValidateFields Then FindRecordsAndPopulateGrid
        Case 1
            SeekAndEditRecord _
                grdSalesIncomingVehicles.CellText(grdSalesIncomingVehicles.CurRow, "InvoiceTrnID"), _
                grdSalesIncomingVehicles.CellText(grdSalesIncomingVehicles.CurRow, "WindowTitle"), _
                IIf(txtRefersTo.text = "2", "Καρτέλα πελάτη", "Καρτέλα προμηθευτή"), _
                txtTable.text, _
                txtRefersTo.text
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

Private Function ValidateFields()

    ValidateFields = False
    
    'Από <= Εως
    If DisplayMessage(14, 4, 1, "", mskIssueFrom.text, mskIssueTo.text) Then mskIssueFrom.SetFocus: Exit Function
    
    ValidateFields = True

End Function

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function
    
    If Not blnStatus Then
        ClearFields grdSalesIncomingVehicles, lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
        frmCriteria(0).Visible = True
        mskIssueFrom.SetFocus
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
    With grdSalesIncomingVehicles
        .Clear
        .Editable = False
        .Redraw = False
        .RowMode = False
    End With
    
    'Πωλήσεις
    strSQL = "SELECT InvoiceID, InvoiceRefersToID, InvoiceIssueDate, Customers.Description, Customers.Address, Codes.CodeShortDescription, InvoicePlates, InvoiceNo, InvoiceGrossAmount, InvoiceInTime, InvoiceTrnID " _
    & "FROM ((Invoices " _
    & "INNER JOIN " & txtTable.text & " ON Invoices.InvoicePersonID = " & txtTable.text & ".ID) " _
    & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID) "
    
    'Τύπος κίνησης
    strThisParameter = "intInvoiceID Integer"
    strThisQuery = "Invoices.InvoiceRefersToID = intInvoiceID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtRefersTo.text)
    
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
        
    'Πελάτης
    If txtPersonID.text <> "" Then
        strThisParameter = "intPerson Integer"
        strThisQuery = "Invoices.InvoicePersonID = intPerson"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtPersonID.text)
    End If
    
    'Αρ. κυκλοφορίας
    If txtPlates.text <> "" Then
        If Left(txtPlates.text, 1) <> "*" Then
            strThisParameter = "strPlates String"
            strThisQuery = "Left(Invoices!InvoicePlates,Len(strPlates)) = strPlates"
            strLogic = " AND "
            GoSub UpdateSQLString
            arrQuery(intIndex) = txtPlates.text
        End If
        If Left(txtPlates.text, 1) = "*" Then
            strThisParameter = "strPlates String"
            strThisQuery = "InStr(Invoices!InvoicePlates, " & "'" & Right(txtPlates.text, Len(txtPlates.text) - 1) & "'" & ") "
            strLogic = " AND "
            GoSub UpdateSQLString
            arrQuery(intIndex) = txtPlates.text
        End If
    End If
    
    'Ταξινόμηση
    strOrder = " ORDER BY InvoiceIssueDate, InvoiceInTime"
    
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
    UpdateButtons Me, 5, 0, 0, 0, 0, 1, 0
    cmdButton(4).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        Do While Not .EOF
            grdSalesIncomingVehicles.AddRow
            lngRow = grdSalesIncomingVehicles.RowCount
            grdSalesIncomingVehicles.CellValue(lngRow, "AA") = lngRowCount + 1
            grdSalesIncomingVehicles.CellValue(lngRow, "InvoiceID") = !InvoiceID
            grdSalesIncomingVehicles.CellValue(lngRow, "InvoiceIssueDate") = !InvoiceIssueDate
            grdSalesIncomingVehicles.CellValue(lngRow, "PersonDescription") = !Description
            grdSalesIncomingVehicles.CellValue(lngRow, "Plates") = !InvoicePlates
            grdSalesIncomingVehicles.CellValue(lngRow, "Address") = !Address
            grdSalesIncomingVehicles.CellValue(lngRow, "InvoiceGrossAmount") = !InvoiceGrossAmount
            grdSalesIncomingVehicles.CellValue(lngRow, "CodeShortDescription") = !CodeShortDescription
            grdSalesIncomingVehicles.CellValue(lngRow, "InvoiceNo") = !InvoiceNo
            grdSalesIncomingVehicles.CellValue(lngRow, "WindowTitle") = UpdateWindowTitle(!InvoiceRefersToID)
            grdSalesIncomingVehicles.CellValue(lngRow, "PersonTableName") = UpdateTableName(!InvoiceRefersToID)
            grdSalesIncomingVehicles.CellValue(lngRow, "InvoiceTrnID") = !InvoiceTrnID
            grdSalesIncomingVehicles.CellValue(lngRow, "InvoiceRefersToID") = !InvoiceRefersToID
            grdSalesIncomingVehicles.CellValue(lngRow, "InvoiceInTime") = !InvoiceInTime
            '///
            FillArray curGrandTotal, grdSalesIncomingVehicles.CellValue(lngRow, "InvoiceGrossAmount")
            '///
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
        ClearFields grdSalesIncomingVehicles
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
    ClearFields grdSalesIncomingVehicles, frmProgress
    cmdButton(4).Caption = "Νέα αναζήτηση"
    DisplayErrorMessage True, Err.Description
    
End Function

Private Sub cmdIndex_Click(Index As Integer)

    Dim strCategoryCriteria As String
    
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case Index
        Case 0
            'Πελάτης
            If txtPersonDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtPersonDescription.text, txtTable.text, "Description", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 3, 0, 1, 2, "ID", "Περιγραφή", "Α.Φ.Μ.", 0, 50, 15, 1, 0, 1)
            txtPersonID.text = tmpTableData.strCode
            txtPersonDescription.text = tmpTableData.strOneField
    End Select

End Sub

Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdSalesIncomingVehicles, 44, GetSetting(strAppTitle, "Layout Strings", "grdSalesIncomingVehicles"), _
            "05NCNAA,05NCNInvoiceID,05NCNInvoiceTrnID,05NCNInvoiceRefersToID,10NLNWindowTitle,10NCNPersonTableName,10NCDXInvoiceIssueDate,50NLNPersonDescription,40NLNAddress,20NLNPlates,40NCNCodeShortDescription,10NCNXInvoiceNo,10NRFInvoiceGrossAmount,10NCTXInvoiceInTime,03NCNSelected", _
            "A/A,InvoiceID,InvoiceTrnID,InvoiceRefersToID,Παράθυρο,Πίνακας,Ημερομηνία έκδοσης,Συναλλασόμενος,Διεύθυνση,Αρ. κυκλοφορίας,Παραστατικό,Νο παραστατικού,Συνολική αξία,Ωρα εξόδου,Ε"
        Me.Refresh
        frmCriteria(0).Visible = True
        mskIssueFrom.SetFocus
    End If
    
    'AddDummyLines grdSalesIncomingVehicles, 5, 5, 5, 5, 5, 5, 12, 50, 40, 20, 4, 10, 10, 5, 4
    
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

    SetUpGrid lstIconList, grdSalesIncomingVehicles
    PositionControls Me, True, grdSalesIncomingVehicles
    ColorizeControls Me, True
    ClearFields mskIssueFrom, mskIssueTo, txtPersonID, txtPersonDescription, txtPlates, lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
    UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1

End Sub

Private Sub grdSalesIncomingVehicles_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    If grdSalesIncomingVehicles.RowCount = 0 Then Exit Sub
    
    grdSalesIncomingVehicles.RemoveRow (grdSalesIncomingVehicles.RowCount): grdSalesIncomingVehicles.RemoveRow (grdSalesIncomingVehicles.RowCount)

End Sub

Private Sub grdSalesIncomingVehicles_ColHeaderMouseEnter(ByVal lCol As Long)

    grdSalesIncomingVehicles.Header.Buttons = True

End Sub

Private Sub grdSalesIncomingVehicles_ColHeaderMouseLeave(ByVal lCol As Long)

    grdSalesIncomingVehicles.Header.Buttons = False
    
End Sub

Private Sub grdSalesIncomingVehicles_ContentsSorted()

    AddGridRowWithTotals grdSalesIncomingVehicles, 0, "PersonDescription", False, strMessages(32), True, curGrandTotal(), 1, 2, 0, "InvoiceGrossAmount"
    
End Sub

Private Sub grdSalesIncomingVehicles_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)

    cmdButton(1).Enabled = CheckToEnableButton(grdSalesIncomingVehicles, lRow, "InvoiceID")

End Sub

Private Sub grdSalesIncomingVehicles_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1
    
End Sub

Private Sub grdSalesIncomingVehicles_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdSalesIncomingVehicles_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeyInsert Or KeyCode = vbKeyDelete Or KeyCode = vbKeySpace Then
        grdSalesIncomingVehicles.CellIcon(grdSalesIncomingVehicles.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdSalesIncomingVehicles, KeyCode, grdSalesIncomingVehicles.CurRow, "InvoiceID"))
        lblSelectedGridLines.Caption = CountSelected(grdSalesIncomingVehicles)
        lblSelectedGridTotals.Caption = SumSelectedGridRows(grdSalesIncomingVehicles, False, "InvoiceGrossAmount")
    End If

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdSalesIncomingVehicles", grdSalesIncomingVehicles.LayoutCol

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

