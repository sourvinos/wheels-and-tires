VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form PersonsLedger 
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
      Left            =   8925
      TabIndex        =   38
      Top             =   7650
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "PersonsLedger.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "PersonsLedger.frx":001C
         BarPictureMode  =   0
         BackPictureMode =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         TabIndex        =   40
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
         Height          =   2340
         Index           =   1
         Left            =   150
         TabIndex        =   42
         Top             =   2625
         Visible         =   0   'False
         Width           =   8715
         Begin VB.Frame Frame1 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Height          =   540
            Left            =   1875
            TabIndex        =   47
            Top             =   1725
            Width           =   4890
            Begin GurhanButtonOCX.GurhanButton cmdButton 
               Height          =   390
               Index           =   9
               Left            =   2475
               TabIndex        =   49
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
               Index           =   8
               Left            =   225
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   75
               Width           =   2190
               _ExtentX        =   3863
               _ExtentY        =   688
               Caption         =   "Αποστολή"
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
         Begin UserControls.newText txtEmail 
            Height          =   465
            Left            =   2100
            TabIndex        =   46
            Top             =   825
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
            Index           =   5
            Left            =   0
            TabIndex        =   52
            Top             =   1650
            Width           =   8715
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H000000C0&
            Caption         =   "Email παραλήπτη"
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
            Height          =   255
            Index           =   1
            Left            =   450
            TabIndex        =   45
            Top             =   900
            Width           =   1215
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
            Left            =   4050
            TabIndex        =   44
            Top             =   75
            Width           =   4515
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   1440
            Index           =   3
            Left            =   1650
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
            Left            =   8250
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
            Index           =   5
            Left            =   0
            Top             =   750
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808000&
            Caption         =   "Αποστολή με email"
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
            TabIndex        =   43
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
            Height          =   540
            Index           =   1
            Left            =   0
            TabIndex        =   51
            Top             =   0
            Width           =   8715
         End
      End
      Begin VB.Frame frmButtonFrame 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   75
         TabIndex        =   25
         Top             =   8850
         Width           =   11790
         Begin GurhanButtonOCX.GurhanButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   26
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
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            Caption         =   "Επεξεργασία εγγραφής"
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
            TabIndex        =   28
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
            TabIndex        =   29
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
            TabIndex        =   30
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
            TabIndex        =   31
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
            Index           =   5
            Left            =   7350
            TabIndex        =   41
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
         Begin GurhanButtonOCX.GurhanButton cmdButton 
            Height          =   690
            Index           =   4
            Left            =   5925
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            Caption         =   "Δημιουργία αρχείου XLS"
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
         Height          =   2190
         Left            =   8850
         TabIndex        =   16
         Tag             =   "Hidden"
         Top             =   5325
         Visible         =   0   'False
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
            TabIndex        =   37
            TabStop         =   0   'False
            Text            =   "3"
            Top             =   825
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
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "OppositeTable"
            Top             =   825
            Width           =   1965
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
            Top             =   450
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
            TabIndex        =   34
            TabStop         =   0   'False
            Text            =   "2"
            Top             =   450
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
            TabIndex        =   33
            TabStop         =   0   'False
            Text            =   "RefersTo"
            Top             =   1200
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
            TabIndex        =   32
            TabStop         =   0   'False
            Text            =   "4"
            Top             =   1200
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
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   "Persons.PersonID"
            Top             =   75
            Width           =   1965
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
            Left            =   2100
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   "1"
            Top             =   75
            Width           =   2340
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   1575
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   2296
            Images          =   "PersonsLedger.frx":0038
            Version         =   131072
            KeyCount        =   2
            Keys            =   "ÿ"
         End
      End
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   3690
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   5025
         Width           =   8640
         Begin VB.CheckBox chkCriteriaChecksAnalysis 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Caption         =   "Ανάλυση αξιογράφων"
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
            Left            =   1575
            TabIndex        =   5
            Top             =   2175
            Width           =   4065
         End
         Begin VB.CheckBox chkCriteriaOnlyQty 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Caption         =   "Μόνο ποσότητες"
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
            Left            =   1575
            TabIndex        =   7
            Top             =   2775
            Width           =   4065
         End
         Begin VB.CheckBox chkCriteriaZeroInvoices 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Caption         =   "Να συμπεριληφθούν τα μηδενικής αξίας παραστατικά"
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
            Left            =   1575
            TabIndex        =   6
            Top             =   2475
            Width           =   4065
         End
         Begin VB.CheckBox chkCriteriaItemsAnalysis 
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
            Left            =   1575
            TabIndex        =   4
            Top             =   1875
            Width           =   4065
         End
         Begin UserControls.newText txtPersonDescription 
            Height          =   465
            Left            =   1575
            TabIndex        =   1
            Top             =   825
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
         Begin UserControls.newDate mskIssueFrom 
            Height          =   465
            Left            =   1575
            TabIndex        =   2
            Top             =   1350
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
         Begin UserControls.newDate mskIssueTo 
            Height          =   465
            Left            =   3150
            TabIndex        =   3
            Top             =   1350
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
            Index           =   0
            Left            =   7800
            TabIndex        =   21
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
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            PicNormal       =   "PersonsLedger.frx":0950
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin VB.Label lblLabel 
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
            Index           =   0
            Left            =   450
            TabIndex        =   20
            Top             =   900
            Width           =   690
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
            Left            =   3975
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
            Left            =   1125
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
            Left            =   8175
            Top             =   900
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
            Height          =   690
            Index           =   4
            Left            =   0
            TabIndex        =   15
            Top             =   3225
            Width           =   8640
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
            BackColor       =   &H000080FF&
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
            Top             =   1425
            Width           =   690
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
            Width           =   8640
         End
      End
      Begin iGrid300_10Tec.iGrid grdPersonsLedger 
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
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   315
         Left            =   2550
         TabIndex        =   24
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
         Caption         =   "Καρτέλα συναλλασόμενου"
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
         TabIndex        =   12
         Top             =   75
         Width           =   5730
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
Attribute VB_Name = "PersonsLedger"
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
Dim blnPeriodIsGiven As Boolean

Dim curPrevious() As Currency
Dim curPeriod() As Currency

Private Function CalculatePreviousPeriod(myRecordset As Recordset)

    If IsDate(mskIssueFrom.text) Then
        With myRecordset
            Do While !InvoiceIssueDate < CDate(mskIssueFrom.text)
                FillArray curPrevious, _
                    CalculateDebitCreditAndBalance("Debit", "Persons", !InvoiceGrossAmount, !CodeCustomers, !CodeSuppliers, "", !PaymentWayCreditID, !InvoiceRefersToID), _
                    CalculateDebitCreditAndBalance("Credit", "Persons", !InvoiceGrossAmount, !CodeCustomers, !CodeSuppliers, "", !PaymentWayCreditID, !InvoiceRefersToID)
                UpdateProgressBar Me
                .MoveNext
                DoEvents
                If .EOF Then
                    Exit Do
                Else
                    If Not blnProcessing Then Exit Function
                    If !InvoiceIssueDate >= CDate(mskIssueFrom.text) Then
                        Exit Do
                    End If
                End If
            Loop
            curPrevious(2) = curPrevious(0) - curPrevious(1)
            curPrevious(3) = 0
            CalculatePreviousPeriod = curPrevious()
        End With
    End If

End Function

Private Function DisplayEmailFrame()

    frmCriteria(1).Visible = True
    EnableFields txtEmail
    txtEmail.SetFocus
    UpdateButtons Me, 9, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1

End Function

Private Function HideEmailFrame()

    frmCriteria(1).Visible = False
    UpdateButtons Me, 9, 0, IIf(CheckForLoadedForm("PersonsTransactions,CommonTransactions"), 0, 1), 1, 1, 1, IIf(chkCriteriaOnlyQty.Value = 1, 1, 0), 1, 0, 0, 0
    grdPersonsLedger.SetFocus

End Function


Private Function SeekAndEditRecord(myInvoiceTrnID, myWindowTitle, myTable, myNextWindowTitle, myRefersTo, myOppositeTable)
    
    Dim blnFound As Boolean
    
    Select Case myRefersTo
        Case "1", "2"
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
        Case "3", "4"
            blnFound = Not SimpleSeek("Invoices", "TrnID", Val(myInvoiceTrnID))
            If blnFound Then
                PersonsTransactions.DoSharedStuff myInvoiceTrnID, myWindowTitle, myTable, myRefersTo, myOppositeTable
            Else
                DisplayMessage 17, 4, 1, ""
            End If
    End Select

End Function

Private Function FindRecordsAndPopulateGrid()

    Dim blnEnableEdit As Boolean
    
    If RefreshList > 0 Then
        UpdateRecordCount lblRecordCount, lngRowCount
        UpdateCriteriaLabels txtPersonDescription.text, mskIssueFrom.text, mskIssueTo.text
        If blnPeriodIsGiven Then
            AddGridRowWithTotals grdPersonsLedger, chkCriteriaOnlyQty.Value, "CodeDescription", strMessages(36), curPeriod(), 3, 2, 0, "Debit", "Credit", "Balance", "Qty"
            ColorizeCells grdPersonsLedger, grdPersonsLedger.RowCount, "Debit", "Credit", "Balance"
        End If
        If chkCriteriaOnlyQty.Value = 0 Or (Not blnPeriodIsGiven And chkCriteriaOnlyQty = 1) Then
            AddGridRowWithTotals grdPersonsLedger, chkCriteriaOnlyQty.Value, "CodeDescription", strMessages(32), curGrandTotal(), 3, IIf(blnPeriodIsGiven, 1, 2), 0, "Debit", "Credit", "Balance", "Qty"
            ColorizeCells grdPersonsLedger, grdPersonsLedger.RowCount, "Debit", "Credit", "Balance"
        End If
        EnableGrid grdPersonsLedger, False
        HighlightRow grdPersonsLedger, 1, "", True
        blnEnableEdit = CheckToEnableButton(grdPersonsLedger, 1, "AA")
        UpdateButtons Me, 9, 0, IIf(CheckForLoadedForm("PersonsTransactions,CommonTransactions"), 0, blnEnableEdit), 1, 1, 1, IIf(chkCriteriaOnlyQty.Value = 1, 1, 0), 1, 0, 0, 0
    Else
        UpdateButtons Me, 9, 1, 0, 0, 0, 0, 0, 0, 1, 0, 0
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
        If txtPersonDescription.Enabled Then txtPersonDescription.SetFocus Else mskIssueFrom.SetFocus
    End If
    
End Function

Function CreateUnicodeFile(myPrinterType, myEAFDSSString, myInvoiceHeight, myDetailLines, myTopMargin, myLeftMargin)

    On Error GoTo ErrTrap
    
    Dim lngRow As Long
    Dim intProcessedDetailLines As Integer
    
    Dim intPageNo As Integer
    
    intPageNo = 0
    intProcessedDetailLines = 0
    Dim curBalance As Currency
    
    Dim curTotals(3) As Currency
    
    Open strUnicodeFile For Output As #1
    InitReport myPrinterType, myEAFDSSString, myInvoiceHeight
    GoSub Headers
    
    'Εγγραφές
    With grdPersonsLedger
        For lngRow = 1 To .RowCount
            If .CellText(lngRow, "InvoiceTrnID") <> "" Then
                Print #1, Tab(1); .CellText(lngRow, "InvoiceIssueDate"); Tab(12); Left(.CellText(lngRow, "CodeDescription"), 32); Tab(45); .CellText(lngRow, "InvoiceNo"); Tab(52); .CellText(lngRow, "Plates");
            Else
                Print #1, Tab(12); .CellText(lngRow, "CodeDescription");
            End If
            If chkCriteriaOnlyQty.Value = 0 Then Print #1, Tab(82 - Len(.CellText(lngRow, "Debit"))); .CellText(lngRow, "Debit"); Tab(98 - Len(.CellText(lngRow, "Credit"))); .CellText(lngRow, "Credit"); Tab(112 - Len(.CellText(lngRow, "Balance"))); .CellText(lngRow, "Balance")
            If chkCriteriaOnlyQty.Value = 1 Then Print #1, Tab(112 - Len(.CellText(lngRow, "Qty"))); .CellText(lngRow, "Qty")
            '///
            If .CellText(lngRow, "Debit") <> "" And .CellText(lngRow, "Credit") <> "" Then
                curBalance = .CellText(lngRow, "Debit") - .CellText(lngRow, "Credit")
            Else
                curBalance = 0
            End If
            DoRunningTotal curTotals, .CellText(lngRow, "Debit"), .CellText(lngRow, "Credit"), curBalance, Val(.CellText(lngRow, "Qty"))
            '///
            intProcessedDetailLines = intProcessedDetailLines + 1
            If intProcessedDetailLines > Val(myDetailLines) Then
                Print #1, ""
                AddTotalsToOutputFile Space(11) & strMessages(30), curTotals(), IIf(chkCriteriaOnlyQty.Value = 0, "082FY,098FY,112FY", "000IN,000IN,000IN,104IY")
                GoSub Headers
                AddTotalsToOutputFile Space(11) & strMessages(31), curTotals(), IIf(chkCriteriaOnlyQty.Value = 0, "082FY,098FY,112FY", "000IN,000IN,000IN,104IY")
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
    PrintHeadings 111, intPageNo, CustomUpperCase(lblTitle.Caption), CustomUpperCase(strCriteriaA), CustomUpperCase(strCriteriaB), myTopMargin
    If chkCriteriaOnlyQty.Value = 0 Then PrintColumnHeadings 1, "ΗΜΕΡΟΜΗΝΙΑ", 12, "ΠΑΡΑΣΤΑΤΙΚΟ", 45, "ΝΟ", 52, "ΑΡ. ΚΥΚΛΟΦΟΡΙΑΣ", 76, "ΧΡΕΩΣΗ", 90, "ΠΙΣΤΩΣΗ", 104, "ΥΠΟΛΟΙΠΟ"
    If chkCriteriaOnlyQty.Value = 1 Then PrintColumnHeadings 1, "ΗΜΕΡΟΜΗΝΙΑ", 12, "ΠΑΡΑΣΤΑΤΙΚΟ", 45, "ΝΟ", 52, "ΑΡ. ΚΥΚΛΟΦΟΡΙΑΣ", 104, "ΠΟΣΟΤΗΤΑ"
    Print #1, ""
    intProcessedDetailLines = 7
    
    Return
    
ErrTrap:
    If Err.Number = 13 Then Resume Next
    Close #1
    CreateUnicodeFile = "Error"
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function SendEmail()

    PrintRecords Me, "CreatePDF", False, "PrinterPrintsReportsID"

    If SendTheEmail Then
        frmCriteria(1).Visible = False
        UpdateButtons Me, 9, 0, IIf(CheckForLoadedForm("PersonsTransactions,CommonTransactions"), 0, 1), 1, 1, 1, IIf(chkCriteriaOnlyQty.Value = 1, 1, 0), 1, 0, 0, 0
        grdPersonsLedger.SetFocus
        DisplayMessage 10, 1, 1, ""
    End If

End Function


Private Function SendTheEmail()

    'Dim oSmtp As New EASendMailObjLib.Mail
    
    'With oSmtp
    '    .LicenseCode = "TryIt"
    '    .FromAddr = strSender
    '    .AddRecipientEx txtEmail.text, 0
    '    .Subject = "ΚΡΟΤΣΗΣ ΕΠΕ - ΚΑΡΤΕΛΑ ΛΟΓΑΡΙΑΣΜΟΥ"
    '    .BodyText = ""
    '    If .AddAttachment(strReportsPathName & "UnicodeFile.pdf") <> 0 Then
    '        SendTheEmail = False
    '        DisplayErrorMessage True, "Το αρχείο δεν βρέθηκε"
    '    Else
    '        .ServerAddr = strServer
    '        .username = strUserName
    '        .password = strPassword
    '        .SSL_init
    '        If .SendMail() <> 0 Then
    '            SendTheEmail = False
    '            DisplayErrorMessage True, "Το email δεν στάλθηκε"
    '        Else
    '            SendTheEmail = True
    '        End If
    '    End If
    'End With
    
End Function


Private Function UpdateCriteriaLabels(myPerson, myIssueFrom, myIssueTo)

    strCriteriaA = "Επωνυμία " & "[ " & myPerson & " ]"
    strCriteriaB = "Εκδοση από " & IIf(myIssueFrom <> "", "[ " & myIssueFrom & " ]", "[ ΟΛΑ ]") & " έως " & IIf(myIssueTo <> "", "[ " & myIssueTo & " ]", "[ ΟΛΑ ]")
    
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

Private Sub chkCriteriaChecksAnalysis_Click()

    If chkCriteriaChecksAnalysis.Value = 1 Then
        chkCriteriaOnlyQty.Value = 0
    End If
    
End Sub

Private Sub chkCriteriaChecksAnalysis_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckForArrows (KeyCode)

End Sub

Private Sub chkCriteriaChecksAnalysis_KeyPress(KeyAscii As Integer)

    ValidateInput (KeyAscii)

End Sub

Private Sub chkCriteriaItemsAnalysis_Click()

    If chkCriteriaItemsAnalysis.Value = 0 Then
        chkCriteriaOnlyQty.Value = 0
    End If

End Sub

Private Sub chkCriteriaItemsAnalysis_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckForArrows (KeyCode)

End Sub

Private Sub chkCriteriaItemsAnalysis_KeyPress(KeyAscii As Integer)

    ValidateInput (KeyAscii)

End Sub

Private Sub chkCriteriaOnlyQty_Click()

    If chkCriteriaOnlyQty.Value = 1 Then
        chkCriteriaItemsAnalysis.Value = 1
        chkCriteriaChecksAnalysis.Value = 0
    End If
    
    grdPersonsLedger.LayoutCol = GetSetting(strAppTitle, "Layout Strings", "grdPersonsLedger" & txtTable.text & IIf(chkCriteriaOnlyQty.Value = 1, "OnlyQty", ""))
    
End Sub

Private Sub chkCriteriaOnlyQty_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckForArrows (KeyCode)

End Sub

Private Sub chkCriteriaOnlyQty_KeyPress(KeyAscii As Integer)

    ValidateInput (KeyAscii)

End Sub

Private Sub chkCriteriaZeroInvoices_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckForArrows (KeyCode)

End Sub

Private Sub chkCriteriaZeroInvoices_KeyPress(KeyAscii As Integer)

    ValidateInput (KeyAscii)

End Sub

Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
        Case 0
            If ValidateFields Then FindRecordsAndPopulateGrid
        Case 1
            SeekAndEditRecord _
                grdPersonsLedger.CellText(grdPersonsLedger.CurRow, "InvoiceTrnID"), _
                grdPersonsLedger.CellText(grdPersonsLedger.CurRow, "WindowTitle"), _
                txtTable.text, _
                "NoNeed", _
                grdPersonsLedger.CellText(grdPersonsLedger.CurRow, "InvoiceRefersToID"), _
                txtOppositeTable.text
        Case 2
            PrintRecords Me, "Print", False, "PrinterPrintsReportsID"
        Case 3
            PrintRecords Me, "CreatePDF", True, "PrinterPrintsReportsID"
        Case 4
            ExportToExcel
        Case 5
            EditGrid
        Case 6
            AbortProcedure False
        Case 7
            AbortProcedure True
        Case 8
            SendEmail
        Case 9
            AbortProcedure False
    End Select
    
End Sub

Private Function ExportToExcel()

    On Error GoTo ErrTrap
    
    Dim blnEnableEdit As Boolean
    Dim lngCurrentRow As Long
    
    Dim lngRow As Long
    Dim lngCol As Long
    Dim xlsRowOffsetFromTop As Long
    Dim xlsColCount As Long
    
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object

    lngCurrentRow = grdPersonsLedger.CurRow
    
    Set oExcel = CreateObject("Excel.Application")
    Set oBook = oExcel.Workbooks.Add
    Set oSheet = oBook.Worksheets(1)
    
    xlsRowOffsetFromTop = 10
    xlsColCount = 5
    
    With oSheet
       
        SetFontNameAndSize oSheet, "Ubuntu Condensed", 11
        AddCompanyData oSheet, xlsColCount
        AddTitle oSheet, lblTitle.Caption, xlsColCount
        AddCriteria oSheet, lblCriteria.Caption, xlsColCount
        AddHeaders oSheet, grdPersonsLedger, xlsColCount, _
            "A", "InvoiceIssueDate", _
            "B", "CodeDescription", _
            "C", "InvoiceNo", _
            "D", "Plates", _
            "E", "Qty"
        AdjustColumnWidths oSheet, "A", 10, "B", 50, "C", 12, "D", 12, "E", 12
                
        InitializeProgressBar Me, strAppTitle, grdPersonsLedger.RowCount
        
        'Προσωρινά
        UpdateButtons Me, 6, 0, 0, 0, 0, 0, 1, 0
        cmdButton(5).Caption = "Διακοπή επεξεργασίας"
        blnProcessing = True
                
        For lngRow = 1 To grdPersonsLedger.RowCount
            UpdateProgressBar Me
            .Range("A" & lngRow + xlsRowOffsetFromTop) = grdPersonsLedger.CellValue(lngRow, "InvoiceIssueDate")
            .Range("B" & lngRow + xlsRowOffsetFromTop) = grdPersonsLedger.CellValue(lngRow, "CodeDescription")
            .Range("C" & lngRow + xlsRowOffsetFromTop) = grdPersonsLedger.CellValue(lngRow, "InvoiceNo")
            .Range("D" & lngRow + xlsRowOffsetFromTop) = grdPersonsLedger.CellValue(lngRow, "Plates")
            .Range("E" & lngRow + xlsRowOffsetFromTop) = grdPersonsLedger.CellText(lngRow, "Qty")
            DoEvents
            If Not blnProcessing Then Exit For
        Next lngRow
        
        AddNumberFormats oSheet, grdPersonsLedger, "Integers", 5, "E"
        
    End With
    
    If blnProcessing Then
        blnError = False
        GoSub DoFinals
        oBook.SaveAs strReportsPathName & format(Date, "yyyy.mm.dd") & "-" & format(Time, "hh.mm.ss") & ".xlsx"
        oExcel.Quit
        DisplayMessage 10, 1, 1, ""
    Else
        GoSub DoFinals
        DisplayMessage 10, 1, 1, ""
    End If
    
    Exit Function
    
DoFinals:
    blnProcessing = False
    ClearFields frmProgress
    blnEnableEdit = CheckToEnableButton(grdPersonsLedger, lngCurrentRow, "AA")
    UpdateButtons Me, 9, 0, IIf(CheckForLoadedForm("PersonsTransactions,CommonTransactions"), 0, blnEnableEdit), 1, 1, 1, IIf(chkCriteriaOnlyQty.Value = 1, 1, 0), 1, 0, 0, 0
    cmdButton(5).Caption = "Νέα αναζήτηση"
    grdPersonsLedger.SetFocus
    
    Return

ErrTrap:
    oBook.Close False
    oExcel.Quit
    GoSub DoFinals
    If blnError Then
        DisplayErrorMessage True, Err.Description
    End If
    
End Function



Private Function EditGrid()

    If Not blnEditingGrid Then
        blnEditingGrid = True
        UpdateButtons Me, 9, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0
        cmdButton(6).Caption = "Ακυρο"
        EnableGrid grdPersonsLedger, True, grdPersonsLedger.CurRow, 13
    Else
        blnEditingGrid = False
        UpdateButtons Me, 9, 0, IIf(CheckForLoadedForm("PersonsTransactions,CommonTransactions"), 0, 1), 1, 1, 1, IIf(chkCriteriaOnlyQty.Value = 1, 1, 0), 1, 0, 0, 0
        EnableGrid grdPersonsLedger, True, grdPersonsLedger.CurRow, 13
    End If

End Function



Private Function ValidateFields()

    ValidateFields = False
    
    'Συναλλασόμενος
    If DisplayMessage(1, 4, 1, "", txtPersonID.text) Then txtPersonDescription.SetFocus: Exit Function
    
    'Από <= Εως
    If DisplayMessage(14, 4, 1, "", mskIssueFrom.text, mskIssueTo.text) Then mskIssueFrom.SetFocus: Exit Function
    
    ValidateFields = True

End Function

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function
    
    If frmCriteria(1).Visible Then
        HideEmailFrame
        Exit Function
    End If
    
    If Not blnStatus Then
        If Not blnEditingGrid Then
            'Δεν επεξεργάζομαι το πλέγμα - καθαρίζω και περιμένω νέα κριτήρια
            ClearFields grdPersonsLedger, lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
            frmCriteria(0).Visible = True
            If txtPersonDescription.Enabled Then txtPersonDescription.SetFocus Else mskIssueFrom.SetFocus
            UpdateButtons Me, 9, 1, 0, 0, 0, 0, 0, 0, 1, 0, 0
        Else
            'Επεξεργάζομαι το πλέγμα - ακυρώνω την επεξεργασία
            EnableGrid grdPersonsLedger, False, grdPersonsLedger.CurRow
            UpdateButtons Me, 9, 0, 1, 1, 1, 1, IIf(chkCriteriaOnlyQty.Value = 1, 1, 0), 1, 0, 0, 0
            blnEditingGrid = False
        End If
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
    Dim rstChecks As Recordset
    Dim rstItems As Recordset

    'Local μεταβλητές
    Dim lngRow As Long
    Dim lngCol As Long
    Dim blnPreviousPeriodHasBeenDisplayed As Boolean
    Dim blnAskedPeriodHasData As Boolean
    
    'Αρχικές τιμές
    ReDim curPrevious(3)
    ReDim curPeriod(3)
    ReDim curGrandTotal(3)
    
    blnPeriodIsGiven = False
    intIndex = 0
    lngRowCount = 0
    frmCriteria(0).Visible = True
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Πλέγμα
    With grdPersonsLedger
        .Clear
        .Editable = False
        .Redraw = False
        .RowMode = False
    End With
    
    'Αγορές, πωλήσεις, κινήσεις πελατών και προμηθευτών
    If txtRefersTo.text <> "5" Then
        strSQL = "SELECT InvoiceID, InvoiceIssueDate, InvoiceNo, InvoiceRefersToID, InvoiceGrossAmount, InvoiceTrnID, InvoicePersonID, InvoiceInDate, InvoicePlates, PaymentWayDescription, PaymentWayCreditID, CodeDescription, CodeInventoryQty, CodeIsPhysicalThingID, CodeSuppliers, CodeCustomers, DeliveryPointDescription  " _
        & "FROM (((Invoices " _
        & "INNER JOIN " & txtTable.text & " ON Invoices.InvoicePersonID = " & txtTable.text & ".ID) " _
        & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID) " _
        & "INNER JOIN PaymentWays ON Invoices.InvoicePaymentWayID = PaymentWays.PaymentWayID) " _
        & "INNER JOIN DeliveryPoints ON Invoices.InvoiceDeliveryPointID = DeliveryPoints.DeliveryPointID "
    End If
    
    'Αγορές - Πωλήσεις
    strThisParameter = "intInvoiceIDa Integer"
    strThisQuery = "(Invoices.InvoiceRefersToID = intInvoiceIDa"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtRefersTo.text) - 2
    
    'Πληρωμές - Εισπράξεις
    strThisParameter = "intInvoiceIDb Integer"
    strThisQuery = "Invoices.InvoiceRefersToID = intInvoiceIDb)"
    strLogic = " OR "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtRefersTo.text)
    
    'Εως
    If IsDate(mskIssueTo.text) Then
        strThisParameter = "datIssueTo Date"
        strThisQuery = "Invoices.InvoiceIssueDate <= datIssueTo"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskIssueTo.text)
    End If
        
    'Συναλλασόμενος
    strThisParameter = "intPerson Integer"
    strThisQuery = "Invoices.InvoicePersonID = intPerson"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtPersonID.text)
    
    'Παραστατικά μηδενικής αξίας και δεν έχω επιλέξει "μόνο ποσότητες"
    If chkCriteriaZeroInvoices.Value = 0 And chkCriteriaOnlyQty.Value = 0 Then
        strThisParameter = "curInvoiceGrossAmount Currency"
        strThisQuery = "Invoices.InvoiceGrossAmount <> curInvoiceGrossAmount"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = 0
    End If
    
    'Εχω επιλέξει "Μόνο ποσότητες"
    If chkCriteriaOnlyQty.Value = 1 Then
        strThisParameter = "lngCodeIsPhysicalThingID Long"
        strThisQuery = "Codes.CodeIsPhysicalThingID =  lngCodeIsPhysicalThingID"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = 1
    End If
    
    'Ταξινόμηση
    strOrder = " ORDER BY InvoiceIssueDate, InvoiceID, InvoiceCodeID, InvoiceNo"
    
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
    
    'Θα εμφανίσω σύνολα περιόδου;
    If mskIssueFrom.text > "" Then blnPeriodIsGiven = True
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strAppTitle, rstRecordset
    
    'Προσωρινά
    UpdateButtons Me, 9, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0
    cmdButton(6).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        Do While Not .EOF
            If (mskIssueFrom.text <> "" And Not blnPreviousPeriodHasBeenDisplayed) Then
                CalculatePreviousPeriod rstRecordset
                If chkCriteriaOnlyQty.Value = 0 Then
                    AddGridRowWithTotals grdPersonsLedger, chkCriteriaOnlyQty.Value, "CodeDescription", strMessages(31), curPrevious(), 3, 1, 1, "Debit", "Credit", "Balance", "Qty"
                    ColorizeCells grdPersonsLedger, grdPersonsLedger.RowCount - 1, "Debit", "Credit", "Balance"
                    CalculateGrandTotals curPrevious(0), curPrevious(1), curPrevious(2)
                End If
                blnPreviousPeriodHasBeenDisplayed = True
                blnAskedPeriodHasData = False
                If .EOF Then Exit Do
            End If
            If .EOF Then Exit Do
            If chkCriteriaOnlyQty = 0 Or (chkCriteriaOnlyQty = 1 And !CodeIsPhysicalThingID = 1) Then
                grdPersonsLedger.AddRow
                lngRow = grdPersonsLedger.RowCount
                blnAskedPeriodHasData = True
                grdPersonsLedger.CellValue(lngRow, "AA") = lngRowCount + 1
                grdPersonsLedger.CellValue(lngRow, "InvoiceTrnID") = !InvoiceTrnID
                grdPersonsLedger.CellValue(lngRow, "InvoiceRefersToID") = !InvoiceRefersToID
                grdPersonsLedger.CellValue(lngRow, "WindowTitle") = UpdateWindowTitle(!InvoiceRefersToID)
                grdPersonsLedger.CellValue(lngRow, "CodeDescription") = !CodeDescription
                grdPersonsLedger.CellValue(lngRow, "InvoiceNo") = !InvoiceNo
                grdPersonsLedger.CellValue(lngRow, "InvoiceIssueDate") = !InvoiceIssueDate
                grdPersonsLedger.CellValue(lngRow, "Plates") = !InvoicePlates
                grdPersonsLedger.CellValue(lngRow, "DeliveryPointDescription") = !DeliveryPointDescription
                grdPersonsLedger.CellValue(lngRow, "Debit") = CalculateDebitCreditAndBalance("Debit", "Persons", !InvoiceGrossAmount, !CodeCustomers, !CodeSuppliers, "", !PaymentWayCreditID, !InvoiceRefersToID)
                grdPersonsLedger.CellValue(lngRow, "Credit") = CalculateDebitCreditAndBalance("Credit", "Persons", !InvoiceGrossAmount, !CodeCustomers, !CodeSuppliers, "", !PaymentWayCreditID, !InvoiceRefersToID)
                FillArray curPeriod, _
                    grdPersonsLedger.CellValue(lngRow, "Debit"), _
                    grdPersonsLedger.CellValue(lngRow, "Credit"), _
                    grdPersonsLedger.CellValue(lngRow, "Debit") - grdPersonsLedger.CellValue(lngRow, "Credit")
                grdPersonsLedger.CellValue(lngRow, "Balance") = curPrevious(2) + curPeriod(2)
                ColorizeCells grdPersonsLedger, lngRow, "Debit", "Credit", "Balance"
                If chkCriteriaItemsAnalysis.Value = 1 Then GoSub FindItems
                If chkCriteriaChecksAnalysis.Value = 1 Then GoSub FindChecks
                lngRow = lngRow + 1
                lngRowCount = lngRowCount + 1
            End If
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
        ClearFields grdPersonsLedger
    Else
        '///
        CalculateGrandTotals curPeriod(0), curPeriod(1), curPeriod(2), curPeriod(3)
        RefreshList = IIf(blnAskedPeriodHasData, rstRecordset.RecordCount, 0)
        blnProcessing = False
        '///
    End If
    
    'Τελικές ενέργειες
    cmdButton(6).Caption = "Νέα αναζήτηση"
    frmProgress.Visible = False
    frmCriteria(0).Visible = False
    
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
    If Err.Number = 6 Then Err.Description = Err.Description & " ID εγγραφής: " & rstRecordset!InvoiceID
    ClearFields grdPersonsLedger, frmProgress
    cmdButton(6).Caption = "Νέα αναζήτηση"
    DisplayErrorMessage True, Err.Description
        
    Exit Function
    
FindChecks:
    strSQL = "SELECT CheckTrnID, CheckNo, CheckExpireDate, CheckAmount, BankDescription, Description " _
        & "FROM (Checks " _
        & "INNER JOIN Banks ON Checks.CheckBankID = Banks.BankID) " _
        & "LEFT JOIN " & txtTable.text & " ON Checks.CheckIssuedByID = " & txtTable.text & ".ID " _
        & "WHERE CheckTrnID = " & rstRecordset!InvoiceTrnID
    strOrder = " ORDER BY CheckExpireDate"
    TempQuery.SQL = strSQL & strOrder
    Set rstChecks = TempQuery.OpenRecordset()
    With rstChecks
        Do While Not .EOF
            grdPersonsLedger.AddRow
            lngRow = lngRow + 1
            grdPersonsLedger.CellFont(lngRow, "CodeDescription").Name = "Input"
            grdPersonsLedger.CellFont(lngRow, "CodeDescription").Size = "11"
            grdPersonsLedger.CellValue(lngRow, "CodeDescription") = format(!CheckExpireDate, "dd/mm/yyyy")
            grdPersonsLedger.CellValue(lngRow, "CodeDescription") = grdPersonsLedger.CellValue(lngRow, "CodeDescription") & " " & Space(12 - Len(format(!CheckAmount, "#,##0.00"))) & format(!CheckAmount, "#,##0.00")
            grdPersonsLedger.CellValue(lngRow, "CodeDescription") = grdPersonsLedger.CellValue(lngRow, "CodeDescription") & " " & !CheckNo
            grdPersonsLedger.CellValue(lngRow, "CodeDescription") = grdPersonsLedger.CellValue(lngRow, "CodeDescription") & " " & !BankDescription
            For lngCol = 1 To grdPersonsLedger.colCount
                grdPersonsLedger.CellForeColor(lngRow, lngCol) = vbCyan
            Next lngCol
            .MoveNext
        Loop
    End With
    
    Return
    
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
            grdPersonsLedger.AddRow
            lngRow = lngRow + 1
            grdPersonsLedger.CellFont(lngRow, "CodeDescription").Name = "Input"
            grdPersonsLedger.CellFont(lngRow, "CodeDescription").Size = "11"
            grdPersonsLedger.CellValue(lngRow, "Qty") = !Qty
            grdPersonsLedger.CellValue(lngRow, "CodeDescription") = Trim(!ItemDescription) & IIf(!ManufacturerIsShownID = 1, " " & !ManufacturerDescription & " ", " ") & IIf(chkCriteriaOnlyQty.Value = 0, format(!Qty, "#,##0") & IIf(chkCriteriaOnlyQty.Value = 0, " x " & format(!TotalNetPostDiscount / !Qty, "#,##0.00"), "x τεμ"), "")
            For lngCol = 1 To grdPersonsLedger.colCount
                grdPersonsLedger.CellForeColor(lngRow, lngCol) = vbCyan
            Next lngCol
            curPeriod(3) = curPeriod(3) + !Qty
            .MoveNext
        Loop
    End With
    
    Return
    
End Function

Private Sub cmdIndex_Click(Index As Integer)

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case Index
        Case 0
            'Συναλλασόμενος
            If txtPersonDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtPersonDescription.text, txtTable.text, "Description", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 4, 0, 1, 2, 13, "ID", "Περιγραφή", "Α.Φ.Μ.", "Ε", 0, 50, 15, 0, 1, 0, 1, 1, "Persons")
            txtPersonID.text = tmpTableData.strCode
            txtPersonDescription.text = tmpTableData.strOneField
            txtEmail.text = tmpTableData.strThreeField
    End Select

End Sub

Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        frmCriteria(0).Visible = True
        frmCriteria(1).Visible = False
        AddColumnsToGrid grdPersonsLedger, 44, GetSetting(strAppTitle, "Layout Strings", "grdPersonsLedger" & txtTable.text & IIf(chkCriteriaOnlyQty.Value = 1, "OnlyQty", "")), _
            "05NCNAA,05NCNInvoiceTrnID,05NCNInvoiceRefersToID,10NLNWindowTitle,10NCDXInvoiceIssueDate,40NLNCodeDescription,10NCNXInvoiceNo,10NLNPlates,10NRFDebit,10NRFCredit,10NRFBalance,40NLNDeliveryPointDescription,10NRIQty,03NCNSelected", _
            "Α/Α,TrnID,RefersToID,Παράθυρο,Ημερομηνία έκδοσης,Παραστατικό,Νο παραστατικού,Αριθμός κυκλοφορίας,Χρέωση,Πίστωση,Υπόλοιπο,Τόπος παράδοσης,Ποσότητα,Ε"
        Me.Refresh
        If txtPersonDescription.Enabled Then txtPersonDescription.SetFocus Else mskIssueFrom.SetFocus
    End If
    
    'AddDummyLines grdPersonsLedger, 5, 5, 5, 5, 10, 40, 6, 20, 13, 13, 13, 40, 10, 3
    
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
        Case vbKeyM And CtrlDown = 4 And cmdButton(4).Enabled
            cmdButton_Click 4
        Case vbKeyEscape
            If cmdButton(6).Enabled Then cmdButton_Click 6: Exit Function
            If cmdButton(7).Enabled Then cmdButton_Click 7: Exit Function
            If cmdButton(9).Enabled Then cmdButton_Click 9
        Case vbKeyF10 And cmdButton(8).Enabled, vbKeyC And CtrlDown = 4 And cmdButton(8).Enabled
            cmdButton_Click 8
        Case vbKeyF12
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    SetUpGrid lstIconList, grdPersonsLedger
    PositionControls Me, True, grdPersonsLedger
    ColorizeControls Me, True
    ClearFields mskIssueFrom, mskIssueTo, txtPersonID, txtPersonDescription, chkCriteriaChecksAnalysis, chkCriteriaItemsAnalysis, chkCriteriaZeroInvoices, lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals, txtEmail
    DisableFields txtEmail
    UpdateButtons Me, 9, 1, 0, 0, 0, 0, 0, 0, 1, 0, 0

End Sub

Private Sub grdPersonsLedger_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)

    CalculateNewQtyTotal grdPersonsLedger, False

End Sub

Private Function CalculateNewQtyTotal(myGrid As iGrid, myFirstTime As Boolean)

    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngNewTotalQty As Long
    
    Dim intStep As Integer
    Dim lngDelay As Long
    
    For lngRow = 1 To myGrid.RowCount - 2
        lngNewTotalQty = lngNewTotalQty + Val(myGrid.CellValue(lngRow, "Qty"))
    Next lngRow
    
    myGrid.CellValue(myGrid.RowCount, "Qty") = lngNewTotalQty

End Function


Private Sub grdPersonsLedger_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    If grdPersonsLedger.RowCount = 0 Then Exit Sub

    grdPersonsLedger.RemoveRow (grdPersonsLedger.RowCount): grdPersonsLedger.RemoveRow (grdPersonsLedger.RowCount)

End Sub

Private Sub grdPersonsLedger_ColHeaderMouseEnter(ByVal lCol As Long)

    grdPersonsLedger.Header.Buttons = True

End Sub

Private Sub grdPersonsLedger_ColHeaderMouseLeave(ByVal lCol As Long)

    grdPersonsLedger.Header.Buttons = False
    
End Sub

Private Sub grdPersonsLedger_ContentsSorted()

    AddGridRowWithTotals grdPersonsLedger, chkCriteriaOnlyQty.Value, "CodeDescription", strMessages(36), curPeriod(), 3, 2, 0, "Debit", "Credit", "Balance", "Qty"
    
End Sub

Private Sub grdPersonsLedger_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)

    cmdButton(1).Enabled = CheckToEnableButton(grdPersonsLedger, lRow, "AA")
    cmdButton(1).Enabled = IIf(CheckForLoadedForm("PersonsTransactions,CommonTransactions"), 0, cmdButton(1).Enabled)

End Sub

Private Sub grdPersonsLedger_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1
    
End Sub

Private Sub grdPersonsLedger_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdPersonsLedger_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeyInsert Or KeyCode = vbKeyDelete Or KeyCode = vbKeySpace Then
        grdPersonsLedger.CellIcon(grdPersonsLedger.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdPersonsLedger, KeyCode, grdPersonsLedger.CurRow, "InvoiceTrnID"))
        lblSelectedGridLines.Caption = CountSelected(grdPersonsLedger)
        lblSelectedGridTotals.Caption = SumSelectedGridRows(grdPersonsLedger, True, "Debit", "Credit", "Balance")
    End If

End Sub

Private Sub grdPersonsLedger_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid300_10Tec.ETextEditFlags)

    If lCol = 13 Then
        If CheckForAcceptableKey(iKeyAscii) Then
            CaptureNumbers grdPersonsLedger.TextEditText, lRow, lCol, iKeyAscii, True
        Else
            bCancel = True
        End If
    End If

End Sub

Private Sub grdPersonsLedger_TextEditKeyDown(ByVal lRow As Long, ByVal lCol As Long, ByVal KeyCode As Integer, ByVal Shift As Integer)

    If lCol = 13 Then
        If CheckForAcceptableKey(KeyCode) Then
            CaptureNumbers grdPersonsLedger.TextEditText, lRow, lCol, KeyCode, True
        Else
            KeyCode = 0
        End If
    End If
End Sub

Private Sub grdPersonsLedger_TextEditKeyPress(ByVal lRow As Long, ByVal lCol As Long, KeyAscii As Integer)

    If lCol = 13 Then
        If CheckForAcceptableKey(KeyAscii) Then
            CaptureNumbers grdPersonsLedger.TextEditText, lRow, lCol, KeyAscii, True
        Else
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdPersonsLedger" & txtTable.text & IIf(chkCriteriaOnlyQty.Value = 1, "OnlyQty", ""), grdPersonsLedger.LayoutCol

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

