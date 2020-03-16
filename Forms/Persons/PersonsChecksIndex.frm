VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form PersonsChecksIndex 
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
      Left            =   5625
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
         Picture         =   "PersonsChecksIndex.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "PersonsChecksIndex.frx":001C
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
      Begin VB.Frame frmButtonFrame 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   75
         TabIndex        =   23
         Top             =   8850
         Width           =   8940
         Begin GurhanButtonOCX.GurhanButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   24
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
            TabIndex        =   25
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
            TabIndex        =   26
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
            TabIndex        =   27
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
            TabIndex        =   28
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
            TabIndex        =   29
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
         Height          =   2940
         Left            =   5550
         TabIndex        =   14
         Tag             =   "Hidden"
         Top             =   4575
         Visible         =   0   'False
         Width           =   4515
         Begin VB.TextBox txtOppositeRefersTo 
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
            Text            =   "4"
            Top             =   1200
            Width           =   2340
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
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "OppositeRefersTo"
            Top             =   1200
            Width           =   1965
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
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   "HoldedBy"
            Top             =   1950
            Width           =   1965
         End
         Begin VB.TextBox txtHoldedBy 
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
            Text            =   "6"
            Top             =   1950
            Width           =   2340
         End
         Begin VB.TextBox Text3 
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
            Text            =   "IssuedBy"
            Top             =   1575
            Width           =   1965
         End
         Begin VB.TextBox txtIssuedBy 
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
            Text            =   "5"
            Top             =   1575
            Width           =   2340
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
            TabIndex        =   31
            TabStop         =   0   'False
            Text            =   "2"
            Top             =   450
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
            TabIndex        =   30
            TabStop         =   0   'False
            Text            =   "OppositeTable"
            Top             =   450
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
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "Table"
            Top             =   75
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
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "1"
            Top             =   75
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
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "RefersTo"
            Top             =   825
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
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   "3"
            Top             =   825
            Width           =   2340
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   2325
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   2296
            Images          =   "PersonsChecksIndex.frx":0038
            Version         =   131072
            KeyCount        =   2
            Keys            =   "ÿ"
         End
      End
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   2565
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   6150
         Width           =   5340
         Begin UserControls.newDate mskExpireFrom 
            Height          =   465
            Left            =   1800
            TabIndex        =   1
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
         Begin UserControls.newDate mskExpireTo 
            Height          =   465
            Left            =   3375
            TabIndex        =   2
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
         Begin UserControls.newDate mskInFrom 
            Height          =   465
            Left            =   1800
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
         Begin UserControls.newDate mskInTo 
            Height          =   465
            Left            =   3375
            TabIndex        =   4
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
            Left            =   1800
            TabIndex        =   15
            Top             =   75
            Width           =   3315
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   1
            Left            =   1350
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
            Left            =   4875
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
            Height          =   465
            Index           =   4
            Left            =   0
            TabIndex        =   13
            Top             =   2100
            Width           =   5340
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
            TabIndex        =   11
            Top             =   75
            Width           =   1665
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
            TabIndex        =   7
            Top             =   1425
            Width           =   915
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000000C0&
            Caption         =   "Λήξη"
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
            TabIndex        =   12
            Top             =   0
            Width           =   5340
         End
      End
      Begin iGrid300_10Tec.iGrid grdPersonsChecksIndex 
         Height          =   7290
         Left            =   75
         TabIndex        =   8
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
         TabIndex        =   22
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
         Caption         =   "Ημερολόγιο αξιογράφων"
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
         TabIndex        =   10
         Top             =   75
         Width           =   5820
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
         TabIndex        =   9
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
Attribute VB_Name = "PersonsChecksIndex"
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

Private Function SeekAndEditRecord(myInvoiceID, myInvoiceTrnID, myWindowTitle, myNextWindowTitle, myTable, myRefersTo, myOppositeTable, myOppositeRefersTo)
    
    Dim blnFound As Boolean
    
    Select Case txtRefersTo
        Case "3", "4"
            blnFound = Not SimpleSeek("Invoices", "TrnID", myInvoiceTrnID)
            If blnFound Then
                PersonsTransactions.DoSharedStuff myInvoiceTrnID, myWindowTitle, myTable, myRefersTo, myOppositeTable
            Else
                DisplayMessage 17, 4, 1, ""
            End If
    End Select

End Function

Private Function FindRecordsAndPopulateGrid()

    If RefreshList > 0 Then
        UpdateRecordCount lblRecordCount, lngRowCount
        UpdateCriteriaLabels mskExpireFrom.text, mskExpireTo.text, mskInFrom.text, mskInTo.text
        AddGridRowWithTotals grdPersonsChecksIndex, 0, "PersonDescriptionA", False, strMessages(32), True, curGrandTotal(), 1, 2, 0, "CheckAmount"
        EnableGrid grdPersonsChecksIndex, False
        HighlightRow grdPersonsChecksIndex, 1, "", True
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
        mskExpireFrom.SetFocus
    End If
    
End Function

Private Function AddGridRowWithTotals(myGrid As iGrid, myOnlyQty, myMessageColumn, myPrintableLineOrNot, myMessage, myBalance, mySums, myColumnCount, myHowManyBlankLinesBefore, myHowManyBlankLinesAfter, ParamArray myColumns() As Variant)

    Dim intLoop As Integer
    Dim lngRow As Long
    
    myGrid.AddRow , , , , , , , myHowManyBlankLinesBefore
    lngRow = myGrid.RowCount
    
    myGrid.CellValue(lngRow, myMessageColumn) = myMessage
    myGrid.CellValue(lngRow, "LineType") = IIf(myPrintableLineOrNot, "Printable", "")
    
    For intLoop = 0 To myColumnCount - IIf(myBalance, 1, 1)
        myGrid.CellValue(lngRow, myColumns(intLoop)) = IIf(myOnlyQty = 0, mySums(intLoop), 0)
    Next intLoop
    
    If Not myBalance Then 'False μόνο για συγκεντρωτικά καρτέλας
        myGrid.CellValue(lngRow, "Balance") = mySums(0) - mySums(1)
    End If
    
    If myHowManyBlankLinesAfter > 0 Then
        myGrid.AddRow , , , , , , , myHowManyBlankLinesAfter
        myGrid.CellValue(myGrid.RowCount, "LineType") = IIf(myPrintableLineOrNot, "Printable", "")
    End If
    
End Function




Function CreateUnicodeFile(myPrinterType, myEAFDSSString, myInvoiceHeight, myDetailLines, myTopMargin, myLeftMargin)

    On Error GoTo ErrTrap
    
    Dim lngRow As Long
    Dim intProcessedDetailLines As Integer
    Dim intPageNo As Integer
    
    intPageNo = 0
    intProcessedDetailLines = 0
    
    Dim curTotals(0) As Currency
    
    Open strUnicodeFile For Output As #1
    InitReport myPrinterType, myEAFDSSString, myInvoiceHeight
    GoSub Headers
    
    With grdPersonsChecksIndex
        For lngRow = 1 To .RowCount - 2
            Print #1, _
                Tab(1); .CellText(lngRow, "InvoiceIssueDate"); _
                Tab(12); Left(.CellText(lngRow, "PersonDescriptionA"), 40); _
                Tab(53); Left(.CellText(lngRow, "PersonDescriptionB"), 30); _
                Tab(84); Left(.CellText(lngRow, "BankDescription"), 12); _
                Tab(97); .CellText(lngRow, "CheckNo"); _
                Tab(113); .CellText(lngRow, "CheckExpireDate"); _
                Tab(137 - Len(.CellText(lngRow, "CheckAmount"))); .CellText(lngRow, "CheckAmount")
            If .CellText(lngRow, "InvoiceID") <> "" Then DoRunningTotal curTotals, .CellText(lngRow, "CheckAmount")
            intProcessedDetailLines = intProcessedDetailLines + 1
            If intProcessedDetailLines > myDetailLines Then
                If lngRow < .RowCount - 2 Then
                    Print #1, ""
                    AddTotalsToOutputFile Space(11) & strMessages(30), curTotals(), "137FY"
                    GoSub Headers
                    AddTotalsToOutputFile Space(11) & strMessages(31), curTotals(), "137FY"
                    Print #1, ""
                    intProcessedDetailLines = intProcessedDetailLines + 2
                End If
            End If
        Next lngRow
    End With
    
    Print #1, ""
    AddTotalsToOutputFile Space(11) & strMessages(32), curTotals(), "137FY"
    
    Close #1
    
    CreateUnicodeFile = strUnicodeFile
    
    Exit Function
    
Headers:
    intPageNo = intPageNo + 1
    PrintHeadings 136, intPageNo, CustomUpperCase(lblTitle.Caption), CustomUpperCase(strCriteriaA), CustomUpperCase(strCriteriaB), myTopMargin
    PrintColumnHeadings 1, "ΕΚΔΟΣΗ", 12, CustomUpperCase(txtIssuedBy.text), 53, CustomUpperCase(txtHoldedBy.text), 84, "ΤΡΑΠΕΖΑ", 97, "ΝΟ", 113, "ΛΗΞΗ", 133, "ΠΟΣΟ"
    Print #1, ""
    intProcessedDetailLines = 7
    Return

ErrTrap:
    Close #1
    CreateUnicodeFile = "Error"
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function UpdateCriteriaLabels(myIssueFrom, myIssueTo, myInFrom, myInTo)

    strCriteriaA = "Λήξη από " & IIf(myIssueFrom <> "", "[ " & myIssueFrom & " ]", "[ ΟΛΑ ]") & " έως " & IIf(myIssueTo <> "", "[ " & myIssueTo & " ]", "[ ΟΛΑ ]")
    strCriteriaB = " Καταχώρηση από " & IIf(myInFrom <> "", "[ " & myInFrom & " ]", "[ ΟΛΑ ]") & " έως " & IIf(myInTo <> "", "[ " & myInTo & " ]", "[ ΟΛΑ ]")
    lblCriteria.Caption = strCriteriaA & " " & strCriteriaB
    
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


Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
        Case 0
            If ValidateFields Then FindRecordsAndPopulateGrid
        Case 1
            SeekAndEditRecord _
                grdPersonsChecksIndex.CellText(grdPersonsChecksIndex.CurRow, "InvoiceID"), _
                grdPersonsChecksIndex.CellText(grdPersonsChecksIndex.CurRow, "InvoiceTrnID"), _
                grdPersonsChecksIndex.CellText(grdPersonsChecksIndex.CurRow, "WindowTitle"), _
                grdPersonsChecksIndex.CellText(grdPersonsChecksIndex.CurRow, "NextWindowTitle"), _
                txtTable.text, _
                txtRefersTo.text, _
                txtOppositeTable.text, _
                txtOppositeRefersTo.text
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
    
    'Λήξη
    If DisplayMessage(14, 4, 1, "", mskExpireFrom.text, mskExpireTo.text) Then mskExpireFrom.SetFocus: Exit Function
    
    'Καταχώρηση
    If DisplayMessage(14, 4, 1, "", mskInFrom.text, mskInTo.text) Then mskInFrom.SetFocus: Exit Function
    
    ValidateFields = True

End Function

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function
    
    If Not blnStatus Then
        ClearFields grdPersonsChecksIndex, lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
        frmCriteria(0).Visible = True
        mskExpireFrom.SetFocus
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
    Dim tmpRecordset As Recordset
    Dim rstChecks As Recordset
    Dim rstItems As Recordset

    'Local μεταβλητές
    Dim lngRow As Long
    Dim lngCol As Long
    Dim blnPersonsDataFound As Boolean
    
    'Αρχικές τιμές
    ReDim curGrandTotal(3)
    intIndex = 0
    lngRowCount = 0
    frmCriteria(0).Visible = False
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Πλέγμα
    With grdPersonsChecksIndex
        .Clear
        .Editable = False
        .Redraw = False
        .RowMode = False
    End With
    
    'Αξιόγραφα
    strSQL = "SELECT DISTINCT CheckID, BankDescription, CheckNo, CheckExpireDate, CheckAmount, CheckIssuedByID, CheckRefersToID, Invoices.InvoiceIssueDate, Invoices.InvoiceNo, InvoiceCodeID, InvoicePersonID, InvoiceTrnID, InvoiceID, InvoiceRefersToID, InvoiceInDate " _
    & "FROM ((Checks " _
    & "INNER JOIN Banks ON Checks.CheckBankID = Banks.BankID) " _
    & "INNER JOIN Invoices ON Checks.CheckTrnID = Invoices.InvoiceTrnID) "
    
    'Τύπος κίνησης
    strThisParameter = "intRefersTo Integer"
    strThisQuery = "Checks.CheckRefersToID= intRefersTo"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtRefersTo.text)
    
    'Λήξη
    If IsDate(mskExpireFrom.text) Then
        strThisParameter = "datExpireFrom Date"
        strThisQuery = "Checks.CheckExpireDate >= datExpireFrom"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskExpireFrom.text)
    End If
    If IsDate(mskExpireTo.text) Then
        strThisParameter = "datExpireTo Date"
        strThisQuery = "Checks.CheckExpireDate <= datExpireTo"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskExpireTo.text)
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
    
    'Ταξινόμηση
    strOrder = " ORDER BY CheckExpireDate"
    
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
            grdPersonsChecksIndex.AddRow
            lngRow = grdPersonsChecksIndex.RowCount
            grdPersonsChecksIndex.CellValue(lngRow, "AA") = lngRow
            grdPersonsChecksIndex.CellValue(lngRow, "InvoiceID") = !InvoiceID
            grdPersonsChecksIndex.CellValue(lngRow, "InvoiceTrnID") = !InvoiceTrnID
            grdPersonsChecksIndex.CellValue(lngRow, "LineType") = "Printable"
            grdPersonsChecksIndex.CellValue(lngRow, "CheckExpireDate") = !CheckExpireDate
            grdPersonsChecksIndex.CellValue(lngRow, "InvoiceIssueDate") = !InvoiceIssueDate
            grdPersonsChecksIndex.CellValue(lngRow, "CheckNo") = !CheckNo
            grdPersonsChecksIndex.CellValue(lngRow, "PersonDescriptionA") = FindPersonDescription(txtTable.text, "ID", !InvoicePersonID) 'Κάτοχος πληρωτέας ή εκδότης εισπρακτέας
            If !CheckIssuedByID <> 0 Then
                grdPersonsChecksIndex.CellValue(lngRow, "PersonDescriptionB") = FindPersonDescription(txtOppositeTable.text, "ID", !CheckIssuedByID) 'Εκδότης πληρωτέας
            End If
            If txtRefersTo.text = "4" And !CheckNo <> "" Then
                Set tmpRecordset = NewCheckForMatch("CommonDB", "CheckNo, Description", "(Checks", "INNER JOIN Invoices ON Checks.CheckTrnID = Invoices.InvoiceTrnID) INNER JOIN Suppliers ON Invoices.InvoicePersonID = Suppliers.ID", "CheckNo = '" & !CheckNo & "' AND CheckRefersToID = 3 AND CheckExpireDate = #" & !CheckExpireDate & "#", "", "") 'Κάτοχος εισπρακτέας
                If Not tmpRecordset.EOF Then
                    grdPersonsChecksIndex.CellValue(lngRow, "PersonDescriptionB") = tmpRecordset!Description
                End If
            End If
            grdPersonsChecksIndex.CellValue(lngRow, "BankDescription") = !BankDescription
            grdPersonsChecksIndex.CellValue(lngRow, "InvoiceNo") = !InvoiceNo
            grdPersonsChecksIndex.CellValue(lngRow, "CheckAmount") = !CheckAmount
            grdPersonsChecksIndex.CellValue(lngRow, "WindowTitle") = UpdateWindowTitle(!InvoiceRefersToID)
            grdPersonsChecksIndex.CellValue(lngRow, "NextWindowTitle") = UpdateNextWindowTitle(!InvoiceRefersToID)
            '///
            FillArray curGrandTotal, _
                grdPersonsChecksIndex.CellValue(lngRow, "CheckAmount")
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
        ClearFields grdPersonsChecksIndex
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
    ClearFields grdPersonsChecksIndex, frmProgress
    cmdButton(4).Caption = "Νέα αναζήτηση"
    DisplayErrorMessage True, Err.Description
        
    Exit Function
    
FindChecks:
    strSQL = "SELECT CheckTrnID, CheckNo, CheckExpire, CheckAmount, BankDescription, Description " _
        & "FROM (Checks " _
        & "INNER JOIN Banks ON Checks.CheckBankID = Banks.BankID) " _
        & "LEFT JOIN " & txtTable.text & " ON Checks.CheckIssuedByID = " & txtTable.text & ".ID " _
        & "WHERE CheckTrnID = " & rstRecordset!InvoiceTrnID
    strOrder = " ORDER BY CheckExpire"
    TempQuery.SQL = strSQL & strOrder
    Set rstChecks = TempQuery.OpenRecordset()
    With rstChecks
        Do While Not .EOF
            grdPersonsChecksIndex.AddRow
            lngRow = lngRow + 1
            grdPersonsChecksIndex.CellValue(lngRow, "LineType") = "Printable"
            grdPersonsChecksIndex.CellFont(lngRow, "PersonDescription").Name = "Input"
            grdPersonsChecksIndex.CellFont(lngRow, "PersonDescription").Size = "9"
            grdPersonsChecksIndex.CellValue(lngRow, "PersonDescription") = format(!CheckExpire, "dd/mm/yyyy")
            grdPersonsChecksIndex.CellValue(lngRow, "PersonDescription") = grdPersonsChecksIndex.CellValue(lngRow, "PersonDescription") & " " & Space(12 - Len(format(!CheckAmount, "#,##0.00"))) & format(!CheckAmount, "#,##0.00")
            grdPersonsChecksIndex.CellValue(lngRow, "PersonDescription") = grdPersonsChecksIndex.CellValue(lngRow, "PersonDescription") & " " & !CheckNo
            grdPersonsChecksIndex.CellValue(lngRow, "PersonDescription") = grdPersonsChecksIndex.CellValue(lngRow, "PersonDescription") & " " & !BankDescription
            For lngCol = 1 To grdPersonsChecksIndex.colCount
                grdPersonsChecksIndex.CellForeColor(lngRow, lngCol) = vbCyan
            Next lngCol
            .MoveNext
        Loop
    End With
    
End Function

Private Function FindPersonDescription(myTable, myFieldName, myFieldValue)

    Dim tmpRecordset As Recordset

    Set tmpRecordset = NewCheckForMatch("CommonDB", "InvoiceTrnID, Description", "Invoices", "INNER JOIN " & myTable & " ON Invoices.InvoicePersonID = " & myTable & ".ID", myFieldName & " = " & myFieldValue, "", "")
    
    If tmpRecordset.RecordCount = 1 Then FindPersonDescription = tmpRecordset!Description

End Function

Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdPersonsChecksIndex, 44, GetSetting(strAppTitle, "Layout Strings", "grdPersonsChecksIndex"), _
            "06NCIAA,05NCNInvoiceID,05NCNInvoiceTrnID,10NLNWindowTitle,10NLNNextWindowTitle,10NLNLineType,06NCNInvoiceNo,10NCDCheckExpireDate,10NCDInvoiceIssueDate,40NLNBankDescription,40NLNPersonDescriptionA,40NLNPersonDescriptionB,40NCNCheckNo,10NRFCheckAmount,03NCNSelected", _
            "ΑΑ,InvoiceID,InvoiceTrnID,Παράθυρο,Επόμενο παράθυρο,Τύπος γραμμής,Νο παραστατικού,Λήξη,Εκδοση,Τράπεζα," & txtIssuedBy.text & "," & txtHoldedBy.text & ",Νο αξιογράφου,Ποσό,Ε"
        Me.Refresh
        frmCriteria(0).Visible = True
        mskExpireFrom.SetFocus
    End If
    
    'AddDummyLines grdPersonsChecksIndex, 1, 1, 1, 1, 1, 1, 10, 30, 50, 70, 90, 110, 50, 10
    
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

    SetUpGrid lstIconList, grdPersonsChecksIndex
    PositionControls Me, True, grdPersonsChecksIndex
    ColorizeControls Me, True
    ClearFields mskExpireFrom, mskExpireTo, mskInFrom, mskInTo, lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
    UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1

End Sub

Private Sub grdPersonsChecksIndex_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    If grdPersonsChecksIndex.RowCount = 0 Then Exit Sub

    grdPersonsChecksIndex.RemoveRow (grdPersonsChecksIndex.RowCount): grdPersonsChecksIndex.RemoveRow (grdPersonsChecksIndex.RowCount)

End Sub

Private Sub grdPersonsChecksIndex_ColHeaderMouseEnter(ByVal lCol As Long)

    grdPersonsChecksIndex.Header.Buttons = True

End Sub

Private Sub grdPersonsChecksIndex_ColHeaderMouseLeave(ByVal lCol As Long)

    grdPersonsChecksIndex.Header.Buttons = False
    
End Sub

Private Sub grdPersonsChecksIndex_ContentsSorted()

    AddGridRowWithTotals grdPersonsChecksIndex, 0, "PersonDescriptionA", False, strMessages(32), True, curGrandTotal(), 1, 2, 0, "CheckAmount"
    
End Sub

Private Sub grdPersonsChecksIndex_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)

    cmdButton(1).Enabled = CheckToEnableButton(grdPersonsChecksIndex, lRow, "InvoiceID")

End Sub

Private Sub grdPersonsChecksIndex_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1
    
End Sub

Private Sub grdPersonsChecksIndex_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdPersonsChecksIndex_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeyInsert Or KeyCode = vbKeyDelete Or KeyCode = vbKeySpace Then
        grdPersonsChecksIndex.CellIcon(grdPersonsChecksIndex.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdPersonsChecksIndex, KeyCode, grdPersonsChecksIndex.CurRow, "InvoiceID"))
        lblSelectedGridLines.Caption = CountSelected(grdPersonsChecksIndex)
        lblSelectedGridTotals.Caption = SumSelectedGridRows(grdPersonsChecksIndex, False, "CheckAmount")
    End If

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdPersonsChecksIndex", grdPersonsChecksIndex.LayoutCol

End Sub

