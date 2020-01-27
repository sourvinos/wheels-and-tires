VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form CommonTransactionsIndex 
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
      TabIndex        =   21
      Top             =   7650
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "CommonTransactionsIndex.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "CommonTransactionsIndex.frx":001C
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
         TabIndex        =   23
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
         TabIndex        =   47
         Top             =   8850
         Width           =   8940
         Begin GurhanButtonOCX.GurhanButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   48
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
            TabIndex        =   49
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
            TabIndex        =   50
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
            TabIndex        =   51
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
            TabIndex        =   52
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
            TabIndex        =   53
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
         Height          =   4065
         Left            =   9450
         TabIndex        =   27
         Tag             =   "Hidden"
         Top             =   3450
         Visible         =   0   'False
         Width           =   4515
         Begin VB.TextBox txtItemID 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
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
            Text            =   "4"
            Top             =   1200
            Width           =   780
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
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
            Text            =   "Items.ItemID"
            Top             =   1200
            Width           =   3540
         End
         Begin VB.TextBox txtCategoryID 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
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
            TabIndex        =   68
            TabStop         =   0   'False
            Text            =   "3"
            Top             =   825
            Width           =   780
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
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
            Text            =   "Categories.CategoryID"
            Top             =   825
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
            TabIndex        =   61
            TabStop         =   0   'False
            Text            =   "8"
            Top             =   2700
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
            TabIndex        =   60
            TabStop         =   0   'False
            Text            =   "RefersTo"
            Top             =   2700
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
            TabIndex        =   59
            TabStop         =   0   'False
            Text            =   "6"
            Top             =   1950
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
            TabIndex        =   58
            TabStop         =   0   'False
            Text            =   "Table"
            Top             =   1950
            Width           =   3540
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
            TabIndex        =   57
            TabStop         =   0   'False
            Text            =   "OppositeTable"
            Top             =   2325
            Width           =   3540
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
            Left            =   3675
            TabIndex        =   56
            TabStop         =   0   'False
            Text            =   "7"
            Top             =   2325
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
            TabIndex        =   55
            TabStop         =   0   'False
            Text            =   "OppositeRefersTo"
            Top             =   3075
            Width           =   3540
         End
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
            Left            =   3675
            TabIndex        =   54
            TabStop         =   0   'False
            Text            =   "9"
            Top             =   3075
            Width           =   780
         End
         Begin VB.TextBox Text5 
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
            TabIndex        =   44
            TabStop         =   0   'False
            Text            =   "Codes.CodeID"
            Top             =   1575
            Width           =   3540
         End
         Begin VB.TextBox txtCodeID 
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
            TabIndex        =   43
            TabStop         =   0   'False
            Text            =   "5"
            Top             =   1575
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
            TabIndex        =   42
            TabStop         =   0   'False
            Text            =   "DeliveryPoints.DeliveryPointID"
            Top             =   450
            Width           =   3540
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
            TabIndex        =   41
            TabStop         =   0   'False
            Text            =   "2"
            Top             =   450
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
            TabIndex        =   40
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
            TabIndex        =   39
            TabStop         =   0   'False
            Text            =   "1"
            Top             =   75
            Width           =   780
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   3450
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   2296
            Images          =   "CommonTransactionsIndex.frx":0038
            Version         =   131072
            KeyCount        =   2
            Keys            =   "ÿ"
         End
      End
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   6840
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   1875
         Width           =   9240
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
            Left            =   2175
            TabIndex        =   12
            Top             =   5325
            Value           =   1  'Checked
            Width           =   4065
         End
         Begin VB.CheckBox chkCriteriaPrintPersonsData 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Caption         =   "Να εκτυπωθούν τα στοιχεία των συναλλασόμενων"
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
            TabIndex        =   14
            Top             =   5925
            Value           =   1  'Checked
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
            Left            =   2175
            TabIndex        =   13
            Top             =   5625
            Value           =   1  'Checked
            Width           =   5824
         End
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
            TabIndex        =   11
            Top             =   5025
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
         Begin UserControls.newText txtCodeShortDescription 
            Height          =   465
            Left            =   2175
            TabIndex        =   9
            Top             =   3975
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   820
            Alignment       =   2
            ForeColor       =   0
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
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   0
            Left            =   8400
            TabIndex        =   35
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
            PicNormal       =   "CommonTransactionsIndex.frx":0950
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   1
            Left            =   7200
            TabIndex        =   36
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
            PicNormal       =   "CommonTransactionsIndex.frx":0EEA
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   4
            Left            =   3750
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   3975
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
            PicNormal       =   "CommonTransactionsIndex.frx":1484
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newText txtCategoryShortDescription 
            Height          =   465
            Left            =   2175
            TabIndex        =   7
            Top             =   2925
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   820
            Alignment       =   2
            ForeColor       =   0
            MaxLength       =   2
            Text            =   "AA"
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
            Left            =   2850
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   2925
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
            PicNormal       =   "CommonTransactionsIndex.frx":1A1E
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newText txtItemDescription 
            Height          =   465
            Left            =   2175
            TabIndex        =   8
            Top             =   3450
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
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   3
            Left            =   8400
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   3450
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
            PicNormal       =   "CommonTransactionsIndex.frx":1FB8
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newText txtInvoiceNo 
            Height          =   465
            Left            =   2175
            TabIndex        =   10
            Top             =   4500
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   820
            Alignment       =   2
            ForeColor       =   0
            MaxLength       =   7
            Text            =   "999.999"
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
            BackColor       =   &H000000C0&
            Caption         =   "Είδος"
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
            TabIndex        =   66
            Top             =   3525
            Width           =   1290
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
            TabIndex        =   64
            Top             =   3000
            Width           =   1290
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
            Left            =   3300
            TabIndex        =   63
            Top             =   3000
            Width           =   4365
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
            Left            =   4200
            TabIndex        =   38
            Top             =   4050
            Width           =   4365
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
            Index           =   4
            Left            =   450
            TabIndex        =   34
            Top             =   4575
            Width           =   1290
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000000C0&
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
            Index           =   3
            Left            =   450
            TabIndex        =   33
            Top             =   4050
            Width           =   1290
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
            TabIndex        =   32
            Top             =   2475
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
            TabIndex        =   31
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
            TabIndex        =   28
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
            Top             =   1725
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
            TabIndex        =   26
            Top             =   6375
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
            TabIndex        =   24
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
            TabIndex        =   17
            Top             =   1425
            Width           =   1290
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
            TabIndex        =   16
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
            TabIndex        =   25
            Top             =   0
            Width           =   9240
         End
      End
      Begin iGrid300_10Tec.iGrid grdCommonTransactionsIndex 
         Height          =   7290
         Left            =   75
         TabIndex        =   18
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
      Begin iGrid300_10Tec.iGrid grdΣτοιχείαΣυναλλασόμενων 
         Height          =   1515
         Left            =   9450
         TabIndex        =   46
         TabStop         =   0   'False
         Tag             =   "Hidden"
         Top             =   3225
         Visible         =   0   'False
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2672
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
         TabIndex        =   45
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
         TabIndex        =   30
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
         TabIndex        =   29
         Top             =   1125
         Width           =   2565
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ημερολόγιο κινήσεων"
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
         TabIndex        =   20
         Top             =   75
         Width           =   5115
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
         TabIndex        =   19
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
Attribute VB_Name = "CommonTransactionsIndex"
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

Private Function SeekAndEditRecord(myInvoiceTrnID, myWindowTitle, myTable, myRefersTo, myOppositeTable, myOppositeRefersTo)
    
    Dim blnFound As Boolean
    
    Select Case myRefersTo
        Case "1", "2"
            blnFound = Not SimpleSeek("Invoices", "TrnID", myInvoiceTrnID)
            If blnFound Then
                CommonTransactions.DoSharedStuff myInvoiceTrnID, myWindowTitle, myTable, myRefersTo
                If CommonTransactions.Visible Then
                    Unload Me
                    If CommonTransactions.mskInvoiceIssueDate.Enabled Then CommonTransactions.mskInvoiceIssueDate.SetFocus
                Else
                    CommonTransactions.Show 1
                End If
            Else
                DisplayMessage 17, 4, 1, ""
            End If
        Case "3", "4"
            blnFound = Not SimpleSeek("Invoices", "TrnID", myInvoiceTrnID)
            If blnFound Then
                PersonsTransactions.DoSharedStuff myInvoiceTrnID, myWindowTitle, myTable, myRefersTo, myOppositeTable
            Else
                DisplayMessage 17, 4, 1, ""
            End If
        Case "5"
            blnFound = Not SimpleSeek("Invoices", "TrnID", Val(myInvoiceTrnID))
            If blnFound Then
                ItemsTransactions.DoSharedStuff myInvoiceTrnID, myWindowTitle
            Else
                 DisplayMessage 17, 4, 1, ""
            End If
    End Select
    
End Function

Private Function FindRecordsAndPopulateGrid()

    If RefreshList > 0 Then
        UpdateRecordCount lblRecordCount, lngRowCount
        UpdateCriteriaLabels mskIssueFrom.text, mskIssueTo.text, mskInFrom.text, mskInTo.text, txtPersonDescription.text, txtDeliveryPointDescription.text, lblCategoryDescription.Caption, txtItemDescription.text, lblCodeDescription.Caption, txtInvoiceNo.text
        AddGridRowWithTotals grdCommonTransactionsIndex, 0, "PersonDescription", False, strMessages(32), True, curGrandTotal(), 4, 2, 0, "InvoiceRestAmount", "InvoiceVATAmount", "InvoiceGrossAmount", "Qty"
        EnableGrid grdCommonTransactionsIndex, False
        HighlightRow grdCommonTransactionsIndex, 1, "", True
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
    
    If Not myBalance Then 'False μόνο για συγκεντρωτικά καρτέλας
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
    
    Dim curTotals(3) As Currency
    
    Open strUnicodeFile For Output As #1
    InitReport myPrinterType, myEAFDSSString, myInvoiceHeight
    GoSub Headers
    
    With grdCommonTransactionsIndex
        For lngRow = 1 To .RowCount - 2
            Select Case txtRefersTo.text
                Case "1", "2" 'Αγορές - Πωλήσεις
                    If .CellText(lngRow, "InvoiceTrnID") <> "" Then
                        Print #1, _
                        Tab(1); .CellText(lngRow, "InvoiceIssueDate"); _
                        Tab(12); Left(.CellText(lngRow, "PersonDescription"), 30); _
                        Tab(43); Left(.CellText(lngRow, "CodeDescription"), 28); _
                        Tab(72); .CellText(lngRow, "InvoiceNo"); _
                        Tab(79); Left(.CellText(lngRow, "PaymentWayDescription"), 15); _
                        Tab(108 - Len(.CellText(lngRow, "InvoiceRestAmount"))); .CellText(lngRow, "InvoiceRestAmount"); _
                        Tab(122 - Len(.CellText(lngRow, "InvoiceVATAmount"))); .CellText(lngRow, "InvoiceVATAmount"); _
                        Tab(136 - Len(.CellText(lngRow, "InvoiceGrossAmount"))); .CellText(lngRow, "InvoiceGrossAmount")
                        DoRunningTotal curTotals, .CellText(lngRow, "InvoiceRestAmount"), .CellText(lngRow, "InvoiceVATAmount"), .CellText(lngRow, "InvoiceGrossAmount")
                    Else
                        Print #1, Tab(12); .CellText(lngRow, "PersonDescription")
                    End If
                Case "3", "4" 'Προμηθευτές - Πελάτες
                   If .CellText(lngRow, "InvoiceTrnID") <> "" Then
                        Print #1, Tab(0); .CellText(lngRow, "InvoiceIssueDate"); Tab(12); .CellText(lngRow, "PersonDescription"); Tab(63); .CellText(lngRow, "CodeDescription"); Tab(104); .CellText(lngRow, "InvoiceNo"); Tab(125 - Len(.CellText(lngRow, "InvoiceGrossAmount"))); .CellText(lngRow, "InvoiceGrossAmount")
                        DoRunningTotal curTotals, .CellText(lngRow, "InvoiceGrossAmount")
                    Else
                        Print #1, Tab(12); .CellText(lngRow, "PersonDescription"); Tab(125 - Len(.CellText(lngRow, "InvoiceGrossAmount"))); .CellText(lngRow, "InvoiceGrossAmount")
                    End If
                Case "5" 'Είδη
                   If .CellText(lngRow, "InvoiceTrnID") <> "" Then
                        Print #1, _
                            Tab(0); .CellText(lngRow, "InvoiceIssueDate"); _
                            Tab(12); Left(.CellText(lngRow, "CodeDescription"), 15); _
                            Tab(28); .CellText(lngRow, "InvoiceNo"); _
                            Tab(35); Left(.CellText(lngRow, "CategoryDescription"), 20); _
                            Tab(56); .CellText(lngRow, "ItemDescription"); _
                            Tab(107); Left(.CellText(lngRow, "ManufacturerDescription"), 20); _
                            Tab(136 - Len(.CellText(lngRow, "Qty"))); .CellText(lngRow, "Qty")
                        DoRunningTotal curTotals, .CellText(lngRow, "Qty")
                    Else
                        Print #1, Tab(12); .CellText(lngRow, "PersonDescription"); Tab(125 - Len(.CellText(lngRow, "Qty"))); .CellText(lngRow, "Qty")
                    End If
            End Select
            intProcessedDetailLines = intProcessedDetailLines + 1
            If intProcessedDetailLines > myDetailLines Then
                If lngRow < .RowCount - 2 Then
                    Select Case txtRefersTo.text
                        Case "1", "2" 'Αγορές - Πωλήσεις
                            Print #1, ""
                            AddTotalsToOutputFile Space(11) & strMessages(30), curTotals(), "108FY,122FY,136FY"
                            GoSub Headers
                            AddTotalsToOutputFile Space(11) & strMessages(31), curTotals(), "108FY,122FY,136FY"
                            Print #1, '"
                            intProcessedDetailLines = intProcessedDetailLines + 2
                        Case "3", "4" 'Προμηθευτές - πελάτες
                            Print #1, ""
                            AddTotalsToOutputFile Space(11) & strMessages(30), curTotals(), "125FY"
                            GoSub Headers
                            AddTotalsToOutputFile Space(11) & strMessages(31), curTotals(), "125FY"
                            Print #1, '"
                            intProcessedDetailLines = intProcessedDetailLines + 2
                        Case "5" 'Είδη
                            Print #1, ""
                            AddTotalsToOutputFile Space(11) & strMessages(30), curTotals(), "136IY"
                            GoSub Headers
                            AddTotalsToOutputFile Space(11) & strMessages(31), curTotals(), "136IY"
                            Print #1, '"
                            intProcessedDetailLines = intProcessedDetailLines + 2
                    End Select
                End If
            End If
        Next lngRow
    End With
    
    Select Case txtRefersTo.text
        Case "1", "2" 'Αγορές - Πωλήσεις
            Print #1, ""
            AddTotalsToOutputFile Space(11) & strMessages(32), curTotals(), "108FY,122FY,136FY"
        Case "3", "4" 'Προμηθευτές - πελάτες
            Print #1, ""
            AddTotalsToOutputFile Space(11) & strMessages(32), curTotals(), "125FY"
        Case "5" 'Είδη
            Print #1, ""
            AddTotalsToOutputFile Space(11) & strMessages(32), curTotals(), "136IY"
    End Select
    
    'Κενές γραμμές = Αλλαγή σελίδας
    For lngRow = intProcessedDetailLines To myDetailLines
        Print #1, ""
    Next lngRow
    
    'Στοιχεία συναλλασόμενων
    If chkCriteriaPrintPersonsData.Value = 1 Then
        intPageNo = 0
        GoSub PersonsHeaders
        With grdΣτοιχείαΣυναλλασόμενων
            .Sort "Description"
            For lngRow = 1 To .RowCount
                Print #1, Tab(1); Left(.CellText(lngRow, "Description"), 45); Tab(47); Left(.CellText(lngRow, "Address"), 40); Tab(88); .CellText(lngRow, "TaxOfficeDescription"); Tab(119); .CellText(lngRow, "TaxNo")
                intProcessedDetailLines = intProcessedDetailLines + 1
                If intProcessedDetailLines > myDetailLines Then
                    Print #1, ""
                    Print #1, strMessages(24)
                    GoSub PersonsHeaders
                    Print #1, strMessages(13)
                    Print #1, ""
                    intProcessedDetailLines = intProcessedDetailLines + 2
                End If
            Next lngRow
            Print #1, ""
            Print #1, strMessages(25)
        End With
    End If
    
    Close #1
    
    CreateUnicodeFile = strUnicodeFile
    
    Exit Function
    
Headers:
    Select Case txtRefersTo.text
        Case "1", "2" 'Αγορές, πωλήσεις
            intPageNo = intPageNo + 1
            PrintHeadings 135, intPageNo, CustomUpperCase(lblTitle.Caption), CustomUpperCase(strCriteriaA), CustomUpperCase(strCriteriaB), myTopMargin
            PrintColumnHeadings 1, "ΗΜΕΡΟΜΗΝΙΑ", 12, "ΣΥΝΑΛΛΑΣΟΜΕΝΟΣ", 43, "ΠΑΡΑΣΤΑΤΙΚΟ", 72, "ΝΟ", 79, "ΤΡΟΠΟΣ", 102, "ΚΑΘΑΡΗ", 118, "ΑΞΙΑ", 128, "ΣΥΝΟΛΙΚΗ"
            PrintColumnHeadings 78, " ΠΛΗΡΩΜΗΣ", 104, "ΑΞΙΑ", 117, "Φ.Π.Α", 132, "ΑΞΙΑ"
            Print #1, ""
            intProcessedDetailLines = 8
        Case "3", "4" 'Προμηθευτές, πελάτες
            intPageNo = intPageNo + 1
            PrintHeadings 124, intPageNo, CustomUpperCase(lblTitle.Caption), CustomUpperCase(strCriteriaA), CustomUpperCase(strCriteriaB), myTopMargin
            PrintColumnHeadings 1, "ΗΜΕΡΟΜΗΝΙΑ", 12, "ΕΠΩΝΥΜΙΑ", 63, "ΠΑΡΑΣΤΑΤΙΚΟ", 104, "ΝΟ", 121, "ΠΟΣΟ"
            Print #1, ""
            intProcessedDetailLines = 7
        Case "5" 'Είδη
            intPageNo = intPageNo + 1
            PrintHeadings 135, intPageNo, CustomUpperCase(lblTitle.Caption), CustomUpperCase(strCriteriaA), CustomUpperCase(strCriteriaB), myTopMargin
            PrintColumnHeadings 1, "ΗΜΕΡΟΜΗΝΙΑ", 12, "ΠΑΡΑΣΤΑΤΙΚΟ", 28, "ΝO", 35, "ΚΑΤΗΓΟΡΙΑ", 56, "ΕΙΔΟΣ", 107, "ΚΑΤΑΣΚΕΥΑΣΤΗΣ", 128, "ΠΟΣΟΤΗΤΑ"
            Print #1, ""
            intProcessedDetailLines = 7
    End Select
    
    Return
    
PersonsHeaders:
    Dim strReportTitle As String
    intPageNo = intPageNo + 1
    strReportTitle = "Στοιχεία κινηθέντων " & IIf(txtTable.text = "Customers", "πελατών", "προμηθευτών")
    PrintHeadings 135, intPageNo, CustomUpperCase(strReportTitle), CustomUpperCase(strCriteriaA), CustomUpperCase(strCriteriaB), myTopMargin
    PrintColumnHeadings 1, "ΕΠΩΝΥΜΙΑ", 47, "ΔΙΕΥΘΥΝΣΗ", 88, "Δ.Ο.Υ.", 119, "Α.Φ.Μ."
    Print #1, ""
    intProcessedDetailLines = 7
    
    Return
    
ErrTrap:
    Close #1
    CreateUnicodeFile = "Error"
    DisplayErrorMessage True, Err.Description
    
    Return

End Function


Private Function UpdateCriteriaLabels(myIssueFrom, myIssueTo, myInFrom, myInTo, myPerson, myDeliveryPoint, myCategory, myItem, myCodeDescription, myInvoiceNo)

    strCriteriaA = "Εκδοση από " & IIf(myIssueFrom <> "", "[ " & myIssueFrom & " ]", "[ ΟΛΑ ]") & " έως " & IIf(myIssueTo <> "", "[ " & myIssueTo & " ]", "[ ΟΛΑ ]") & " Καταχώρηση από " & IIf(myInFrom <> "", "[ " & myInFrom & " ]", "[ ΟΛΑ ]") & " έως " & IIf(myInTo <> "", "[ " & myInTo & " ]", "[ ΟΛΑ ]")
    
    If txtRefersTo.text <> "5" Then
        strCriteriaB = "Επωνυμία " & IIf(myPerson <> "", "[ " & myPerson & " ]", "[ ΟΛΟΙ ]")
    End If
    
    If txtRefersTo.text = "1" Or txtRefersTo.text = "2" Then
        strCriteriaB = strCriteriaB & " Τόπος παραλαβής " & IIf(myDeliveryPoint <> "", "[ " & myDeliveryPoint & " ]", "[ ΟΛΟΙ ]")
    End If
    
    If txtRefersTo.text = "5" Then
        strCriteriaB = "Κατηγορία " & IIf(myCategory <> "", "[ " & myCategory & " ]", "[ ΟΛΕΣ ]")
        strCriteriaB = strCriteriaB & " Είδος " & IIf(myItem <> "", "[ " & myItem & " ]", "[ ΟΛΑ ]")
    End If
    
    strCriteriaB = strCriteriaB & " Παραστατικό " & IIf(myCodeDescription <> "", "[ " & myCodeDescription & " ]", "[ ΟΛΑ ]")
    strCriteriaB = strCriteriaB & " Νο παραστατικού " & IIf(myInvoiceNo <> "", "[ " & myInvoiceNo & " ]", "[ ΟΛΑ ]")
    
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

Private Sub chkCriteriaChecksAnalysis_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckForArrows (KeyCode)

End Sub

Private Sub chkCriteriaChecksAnalysis_KeyPress(KeyAscii As Integer)

    ValidateInput (KeyAscii)

End Sub

Private Sub chkCriteriaItemAnalysis_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckForArrows (KeyCode)

End Sub

Private Sub chkCriteriaItemAnalysis_KeyPress(KeyAscii As Integer)

    ValidateInput (KeyAscii)

End Sub

Private Sub chkCriteriaPrintPersonsData_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckForArrows (KeyCode)

End Sub

Private Sub chkCriteriaPrintPersonsData_KeyPress(KeyAscii As Integer)

    ValidateInput (KeyAscii)

End Sub

Private Sub chkCriteriaZeroInvoices_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckForArrows (KeyCode)

End Sub

Private Sub chkCriteriaZeroInvoices_KeyPress(KeyAscii As Integer)

    ValidateInput (KeyAscii)

End Sub

Private Sub cmdButton_Click(Index As Integer)

    Dim strWindowTitle As String
    
    Select Case Index
        Case 0
            If ValidateFields Then FindRecordsAndPopulateGrid
        Case 1
            Select Case txtRefersTo.text
                Case "1"
                    strWindowTitle = "Αγορές"
                Case "2"
                    strWindowTitle = "Πωλήσεις"
                Case "3"
                    strWindowTitle = "Κινήσεις προμηθευτών"
                Case "4"
                    strWindowTitle = "Κινήσεις πελατών"
                Case "5"
                    strWindowTitle = "Κινήσεις ειδών"
            End Select
            SeekAndEditRecord _
                grdCommonTransactionsIndex.CellText(grdCommonTransactionsIndex.CurRow, "InvoiceTrnID"), _
                strWindowTitle, _
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
    
    'Εκδοση
    If DisplayMessage(14, 4, 1, "", mskIssueFrom.text, mskIssueTo.text) Then mskIssueFrom.SetFocus: Exit Function
    
    'Καταχώρηση
    If DisplayMessage(14, 4, 1, "", mskInFrom.text, mskInTo.text) Then mskInFrom.SetFocus: Exit Function
    
    ValidateFields = True

End Function

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function
    
    If Not blnStatus Then
        ClearFields grdCommonTransactionsIndex, grdΣτοιχείαΣυναλλασόμενων, frmProgress
        ClearFields lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
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
    With grdCommonTransactionsIndex
        .Clear
        .Editable = False
        .Redraw = False
        .RowMode = False
    End With
    
    'Αγορές, πωλήσεις, κινήσεις πελατών και προμηθευτών
    If txtRefersTo.text <> "5" Then
        strSQL = "SELECT InvoiceID, InvoiceIssueDate, InvoiceNo, InvoiceRefersToID, InvoiceRestAmount, InvoiceVATAmount, InvoiceGrossAmount, InvoiceTrnID, InvoicePersonID, InvoiceInDate, InvoiceExtraChargesAmount, PaymentWayDescription, " & txtTable.text & ".Description, " & txtTable.text & ".TaxNo, " & txtTable.text & ".Address, TaxOfficeDescription, CodeDescription, CodeSuppliers, CodeCustomers, DeliveryPointDescription, CountryShortDescription " _
            & "FROM (((((Invoices " _
            & "INNER JOIN " & txtTable.text & " ON Invoices.InvoicePersonID = " & txtTable.text & ".ID) " _
            & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID) " _
            & "INNER JOIN PaymentWays ON Invoices.InvoicePaymentWayID = PaymentWays.PaymentWayID) " _
            & "INNER JOIN TaxOffices ON " & txtTable.text & ".TaxOfficeID = TaxOffices.TaxOfficeID) " _
            & "INNER JOIN DeliveryPoints ON Invoices.InvoiceDeliveryPointID = DeliveryPoints.DeliveryPointID) " _
            & "INNER JOIN Countries ON " & txtTable.text & ".CountryID = Countries.CountryID "
    End If
    
    'Κινήσεις ειδών
    If txtRefersTo.text = "5" Then
        strSQL = "SELECT InvoiceID, InvoiceIssueDate, InvoiceNo, InvoiceRefersToID, InvoiceRestAmount, InvoiceVATAmount, InvoiceGrossAmount, Invoices.InvoiceTrnID, InvoicePersonID, InvoiceInDate, CodeDescription, Items.ItemDescription, InvoicesTrn.Qty, CodeInventoryQty, CodeInventoryValue, CodeSuppliers, CodeCustomers, ManufacturerDescription, CategoryDescription " _
            & "FROM ((((Invoices " _
            & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID) " _
            & "INNER JOIN InvoicesTrn ON Invoices.InvoiceTrnID = InvoicesTrn.InvoiceTrnID) " _
            & "INNER JOIN Items ON InvoicesTrn.ItemID = Items.ItemID) " _
            & "INNER JOIN Manufacturers ON Items.ItemManufacturerID = Manufacturers.ManufacturerID) " _
            & "INNER JOIN Categories ON Items.ItemCategoryID = Categories.CategoryID "
    End If
    
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
    
    'Τόπος παραλαβής
    If txtDeliveryPointID.text <> "" Then
        strThisParameter = "intDeliveryPointID Integer"
        strThisQuery = "Invoices.InvoiceDeliveryPointID = intDeliveryPointID "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtDeliveryPointID.text)
    End If
    
    'Κατηγορία
    If txtCategoryID.text <> "" Then
        strThisParameter = "intCategoryID Integer"
        strThisQuery = "Items.ItemCategoryID = intCategoryID "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtCategoryID.text)
    End If
    
    'Είδος
    If txtItemID.text <> "" Then
        strThisParameter = "intItemID Integer"
        strThisQuery = "Items.ItemID = intItemID "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtItemID.text)
    End If
    
    'Παραστατικό
    If txtCodeID.text <> "" Then
        strThisParameter = "intCodeID Integer"
        strThisQuery = "Invoices.InvoiceCodeID = intCodeID"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtCodeID.text)
    End If
    
    'Νο παραστατικού
    If txtInvoiceNo.text <> "" Then
        strThisParameter = "intInvoiceNo Integer"
        strThisQuery = "Invoices.InvoiceNo = intInvoiceNo"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtInvoiceNo.text)
    End If
    
    'Παραστατικά μηδενικής αξίας
    If chkCriteriaZeroInvoices.Value = 0 Then
        strThisParameter = "curInvoiceGrossAmount Currency"
        strThisQuery = "Invoices.InvoiceGrossAmount <> curInvoiceGrossAmount"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = 0
    End If
    
    'Ταξινόμηση
    strOrder = " ORDER BY InvoiceIssueDate, InvoiceCodeID, InvoiceNo " & IIf(txtRefersTo.text = "5", ", CategoryDescription, ManufacturerDescription, ItemDescription", "")
    
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
            grdCommonTransactionsIndex.AddRow
            lngRow = grdCommonTransactionsIndex.RowCount
            grdCommonTransactionsIndex.CellValue(lngRow, "AA") = lngRowCount + 1
            grdCommonTransactionsIndex.CellValue(lngRow, "InvoiceID") = !InvoiceID
            grdCommonTransactionsIndex.CellValue(lngRow, "InvoiceTrnID") = !InvoiceTrnID
            grdCommonTransactionsIndex.CellValue(lngRow, "InvoiceIssueDate") = !InvoiceIssueDate
            grdCommonTransactionsIndex.CellValue(lngRow, "InvoiceInDate") = !InvoiceInDate
            If txtRefersTo.text <> "5" Then
                grdCommonTransactionsIndex.CellValue(lngRow, "PersonDescription") = !Description
                grdCommonTransactionsIndex.CellValue(lngRow, "PaymentWayDescription") = !PaymentWayDescription
                grdCommonTransactionsIndex.CellValue(lngRow, "DeliveryPointDescription") = !DeliveryPointDescription
                grdCommonTransactionsIndex.CellValue(lngRow, "InvoiceRestAmount") = IIf((!CodeSuppliers = "-" Or !CodeCustomers = "-") And (txtRefersTo.text = "1" Or txtRefersTo.text = "2"), CCur("-" & !InvoiceRestAmount + !InvoiceExtraChargesAmount), !InvoiceRestAmount + !InvoiceExtraChargesAmount)
                grdCommonTransactionsIndex.CellValue(lngRow, "InvoiceVATAmount") = IIf((!CodeSuppliers = "-" Or !CodeCustomers = "-") And (txtRefersTo.text = "1" Or txtRefersTo.text = "2"), CCur("-" & !InvoiceVATAmount), !InvoiceVATAmount)
                grdCommonTransactionsIndex.CellValue(lngRow, "InvoiceGrossAmount") = IIf((!CodeSuppliers = "-" Or !CodeCustomers = "-") And (txtRefersTo.text = "1" Or txtRefersTo.text = "2"), CCur("-" & IIf(!InvoiceGrossAmount < 0, Abs(!InvoiceGrossAmount), !InvoiceGrossAmount)), !InvoiceGrossAmount)
                grdCommonTransactionsIndex.CellValue(lngRow, "Qty") = 0
            Else
                grdCommonTransactionsIndex.CellValue(lngRow, "CategoryDescription") = !CategoryDescription
                grdCommonTransactionsIndex.CellValue(lngRow, "ItemDescription") = !ItemDescription
                grdCommonTransactionsIndex.CellValue(lngRow, "ManufacturerDescription") = !ManufacturerDescription
                grdCommonTransactionsIndex.CellValue(lngRow, "InvoiceRestAmount") = 0
                grdCommonTransactionsIndex.CellValue(lngRow, "InvoiceVATAmount") = 0
                grdCommonTransactionsIndex.CellValue(lngRow, "InvoiceGrossAmount") = 0
                grdCommonTransactionsIndex.CellValue(lngRow, "Qty") = !Qty
            End If
            grdCommonTransactionsIndex.CellValue(lngRow, "CodeDescription") = !CodeDescription
            grdCommonTransactionsIndex.CellValue(lngRow, "InvoiceNo") = !InvoiceNo
            '///
            FillArray curGrandTotal, _
                grdCommonTransactionsIndex.CellValue(lngRow, "InvoiceRestAmount"), _
                grdCommonTransactionsIndex.CellValue(lngRow, "InvoiceVATAmount"), _
                grdCommonTransactionsIndex.CellValue(lngRow, "InvoiceGrossAmount"), _
                grdCommonTransactionsIndex.CellValue(lngRow, "Qty")
            ColorizeCells grdCommonTransactionsIndex, lngRow, "InvoiceRestAmount", "InvoiceVATAmount", "InvoiceGrossAmount", "Qty"
            '///
            If chkCriteriaItemAnalysis.Value = 1 Then GoSub FindItems
            If chkCriteriaChecksAnalysis.Value = 1 Then GoSub FindChecks
            If chkCriteriaPrintPersonsData.Value = 1 And grdCommonTransactionsIndex.CellText(lngRow, "InvoiceID") <> "" Then GoSub FindPersonsData
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
        ClearFields grdCommonTransactionsIndex, grdΣτοιχείαΣυναλλασόμενων, frmProgress
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
    If Err.Number = 6 Then Err.Description = Err.Description & " ID εγγραφής: " & rstRecordset!InvoiceID
    blnError = True
    ClearFields grdCommonTransactionsIndex, grdΣτοιχείαΣυναλλασόμενων, frmProgress
    cmdButton(4).Caption = "Νέα αναζήτηση"
    DisplayErrorMessage True, Err.Description
        
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
            grdCommonTransactionsIndex.AddRow
            lngRow = lngRow + 1
            grdCommonTransactionsIndex.CellFont(lngRow, "PersonDescription").Name = "Input"
            grdCommonTransactionsIndex.CellFont(lngRow, "PersonDescription").Size = "11"
            grdCommonTransactionsIndex.CellValue(lngRow, "Qty") = !Qty
            grdCommonTransactionsIndex.CellValue(lngRow, "PersonDescription") = Trim(!ItemDescription) & IIf(!ManufacturerIsShownID = 1, " " & !ManufacturerDescription & " ", " ") & format(!Qty, "#,##0") & " x " & format(!TotalNetPostDiscount / !Qty, "#,##0.00") & " = " & format(!TotalNetPostDiscount, "#,##0.00")
            grdCommonTransactionsIndex.CellTextFlags(lngRow, "PersonDescription") = igTextNoClip Or igTextLeft
            For lngCol = 1 To grdCommonTransactionsIndex.colCount
                grdCommonTransactionsIndex.CellForeColor(lngRow, lngCol) = vbCyan
            Next lngCol
            .MoveNext
        Loop
    End With
    
    Return

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
            grdCommonTransactionsIndex.AddRow
            lngRow = lngRow + 1
            grdCommonTransactionsIndex.CellFont(lngRow, "PersonDescription").Name = "Input"
            grdCommonTransactionsIndex.CellFont(lngRow, "PersonDescription").Size = "11"
            grdCommonTransactionsIndex.CellValue(lngRow, "PersonDescription") = format(!CheckExpireDate, "dd/mm/yyyy")
            grdCommonTransactionsIndex.CellValue(lngRow, "PersonDescription") = grdCommonTransactionsIndex.CellValue(lngRow, "PersonDescription") & " " & Space(12 - Len(format(!CheckAmount, "#,##0.00"))) & format(!CheckAmount, "#,##0.00")
            grdCommonTransactionsIndex.CellValue(lngRow, "PersonDescription") = grdCommonTransactionsIndex.CellValue(lngRow, "PersonDescription") & " " & !CheckNo
            grdCommonTransactionsIndex.CellValue(lngRow, "PersonDescription") = grdCommonTransactionsIndex.CellValue(lngRow, "PersonDescription") & " " & !BankDescription
            For lngCol = 1 To grdCommonTransactionsIndex.colCount
                grdCommonTransactionsIndex.CellForeColor(lngRow, lngCol) = vbCyan
            Next lngCol
            .MoveNext
        Loop
    End With
    
    Return
    
FindPersonsData:
    If blnPersonsDataFound = grdΣτοιχείαΣυναλλασόμενων.FindSearchMatchRow("Description", grdCommonTransactionsIndex.CellValue(lngRow, "PersonDescription"), 1) Then
        With grdΣτοιχείαΣυναλλασόμενων
            .AddRow
            .CellValue(.RowCount, "Description") = rstRecordset!Description
            .CellValue(.RowCount, "Address") = rstRecordset!Address
            .CellValue(.RowCount, "TaxOfficeDescription") = rstRecordset!TaxOfficeDescription
            .CellValue(.RowCount, "TaxNo") = rstRecordset!CountryShortDescription & " " & rstRecordset!TaxNo
        End With
    End If
    
    Return
    
End Function

Private Sub cmdIndex_Click(Index As Integer)

    Dim strCategoryCriteria As String
    
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case Index
        Case 0
            'Πελάτης - Προμηθευτής
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
        Case 2
            'Κατηγορία
            If txtCategoryShortDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtCategoryShortDescription.text, "Categories", "CategoryShortDescription", "String", 1, 3)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 3, 0, 1, 2, "ID", "Συντ.", "Περιγραφή", 0, 4, 40, 1, 1, 0)
            txtCategoryID.text = tmpTableData.strCode
            txtCategoryShortDescription.text = tmpTableData.strOneField
            lblCategoryDescription.Caption = tmpTableData.strTwoField
        Case 3
            'Είδος
            If txtItemDescription.text = "" Then Exit Sub
            strCategoryCriteria = IIf(txtCategoryID.text <> "", "AND CategoryID = " & txtCategoryID.text, "")
            Set tmpRecordset = NewCheckForMatch("CommonDB", "ItemID, ItemCategoryID, ItemManufacturerID, CategoryDescription, ManufacturerDescription, ItemDescription, CategoryShortDescription", _
                "((Items", _
                "INNER JOIN Categories ON Items.ItemCategoryID = Categories.CategoryID) " & _
                "INNER JOIN Manufacturers ON Items.ItemManufacturerID = Manufacturers.ManufacturerID) ", _
                "Left(ItemQuickDescription, " & Len(txtItemDescription.text) & ") = '" & txtItemDescription.text & "'" & strCategoryCriteria & "", "", "CategoryDescription, ManufacturerDescription, ItemDescription")
            tmpTableData = DisplayIndex(tmpRecordset, True, True, "Ευρετήριο", 7, 0, 1, 2, 3, 4, 5, 6, "ID", "ID Κατηγορίας", "ID Κατασκευαστή", "Κατηγορία", "Κατασκευαστής", "Περιγραφή", "Συντ. κατηγορίας", 0, 0, 0, 40, 40, 50, 0, 1, 0, 0, 0, 0, 0, 0)
            If tmpTableData.strCode <> "" Then
                txtItemID.text = tmpTableData.strCode
                txtCategoryID.text = tmpTableData.strOneField
                txtCategoryShortDescription.text = tmpTableData.strSixField
                lblCategoryDescription.Caption = tmpTableData.strThreeField
                txtItemDescription.text = tmpTableData.strFiveField
            End If
        Case 4
            'Παραστατικό
            If txtCodeShortDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtCodeShortDescription.text, "Codes", "CodeShortDescription", "String", txtRefersTo.text, 3)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 3, 0, 1, 2, "ID", "Συντ.", "Περιγραφή", 0, 4, 40, 1, 1, 0)
            txtCodeID.text = tmpTableData.strCode
            txtCodeShortDescription.text = tmpTableData.strOneField
            lblCodeDescription.Caption = tmpTableData.strTwoField
    End Select

End Sub

Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        If txtRefersTo.text = "1" Or txtRefersTo.text = "2" Then
            DisableFields txtCategoryShortDescription, txtItemDescription, chkCriteriaChecksAnalysis, cmdIndex(2), cmdIndex(3)
        End If
        If txtRefersTo.text = "2" Then
            DisableFields txtDeliveryPointDescription, cmdIndex(1)
        End If
        If txtRefersTo.text = "3" Or txtRefersTo.text = "4" Then
            DisableFields txtDeliveryPointDescription, txtCategoryShortDescription, txtItemDescription, chkCriteriaItemAnalysis, cmdIndex(1), cmdIndex(2), cmdIndex(3)
        End If
        If txtRefersTo.text = "5" Then
            DisableFields txtPersonDescription, txtDeliveryPointDescription, chkCriteriaItemAnalysis, chkCriteriaChecksAnalysis, chkCriteriaZeroInvoices, chkCriteriaPrintPersonsData, cmdIndex(0), cmdIndex(1)
            chkCriteriaZeroInvoices.Value = 1
        End If
        AddColumnsToGrid grdCommonTransactionsIndex, 44, GetSetting(strAppTitle, "Layout Strings", "grdCommonTransactionsIndex" & txtRefersTo.text), _
            "05NCNAA,05NCNInvoiceID,05NCNInvoiceTrnID,10NCDXInvoiceIssueDate,10NCDXInvoiceInDate,40NLNPersonDescription,40NLNCategoryDescription,40NLNItemDescription,40NLNManufacturerDescription,40NLNPaymentWayDescription,40NLNCodeDescription,10NCNXInvoiceNo,10NRFInvoiceRestAmount,10NRFInvoiceVATAmount,10NRFInvoiceGrossAmount,40NLNDeliveryPointDescription,10NRIQty,03NCNSelected", _
            "A/A,InvoiceID,InvoiceTrnID,Ημερομηνία έκδοσης,Ημερομηνία καταχώρησης,Συναλλασόμενος,Κατηγορία,Είδος,Κατασκευαστής,Τρόπος πληρωμής,Παραστατικό,Νο παραστατικού,Καθαρή αξία,Φ.Π.Α.,Συνολική αξία,Τόπος παραλαβής,Ποσότητα,Ε"
        AddColumnsToGrid grdΣτοιχείαΣυναλλασόμενων, 44, GetSetting(strAppTitle, "Layout Strings", "grdΣτοιχείαΣυναλλασόμενων" & txtTable.text), "50NLNDescription,50NLNAddress,50NLNPhones,40NLNTaxOfficeDescription,10NCNTaxNo", "Επωνυμία,Διεύθυνση,Τηλέφωνα,Οικονομική υπηρεσία,Α.Φ.Μ."
        Me.Refresh
        frmCriteria(0).Visible = True
        mskIssueFrom.SetFocus
    End If
    
    'AddDummyLines grdCommonTransactionsIndex, 5, 5, 5, 5, 5, 10, 10, 10, 10, 50, 30, 50, 30, 40, 40, 10, 10, 10, 10, 40, 6, 4
    'AddDummyLines grdΣτοιχείαΣυναλλασόμενων, "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAA"
    
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
        Case vbKeyF12 And CtrlDown = 4
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    SetUpGrid lstIconList, grdCommonTransactionsIndex, grdΣτοιχείαΣυναλλασόμενων
    PositionControls Me, True, grdCommonTransactionsIndex
    ColorizeControls Me, True
    
    ClearFields mskIssueFrom, mskIssueTo, mskInFrom, mskInTo, txtPersonDescription, txtDeliveryPointDescription, txtCategoryShortDescription, txtItemDescription, txtCodeShortDescription, txtInvoiceNo
    ClearFields chkCriteriaItemAnalysis, chkCriteriaChecksAnalysis, chkCriteriaZeroInvoices, chkCriteriaPrintPersonsData
    ClearFields lblCategoryDescription, lblCodeDescription
    ClearFields txtPersonID, txtDeliveryPointID, txtCategoryID, txtItemID, txtCodeID
    ClearFields lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
    
    UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1

End Sub

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



Private Sub grdCommonTransactionsIndex_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    If grdCommonTransactionsIndex.RowCount = 0 Then Exit Sub
    
    grdCommonTransactionsIndex.RemoveRow (grdCommonTransactionsIndex.RowCount): grdCommonTransactionsIndex.RemoveRow (grdCommonTransactionsIndex.RowCount)

End Sub

Private Sub grdCommonTransactionsIndex_ColHeaderMouseEnter(ByVal lCol As Long)

    grdCommonTransactionsIndex.Header.Buttons = True

End Sub

Private Sub grdCommonTransactionsIndex_ColHeaderMouseLeave(ByVal lCol As Long)

    grdCommonTransactionsIndex.Header.Buttons = False
    
End Sub

Private Sub grdCommonTransactionsIndex_ContentsSorted()

    AddGridRowWithTotals grdCommonTransactionsIndex, 0, "CodeDescription", False, strMessages(32), True, curGrandTotal(), 4, 2, 0, "InvoiceRestAmount", "InvoiceVATAmount", "InvoiceGrossAmount", "Qty"
    
End Sub

Private Sub grdCommonTransactionsIndex_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)

    cmdButton(1).Enabled = CheckToEnableButton(grdCommonTransactionsIndex, lRow, "InvoiceID")

End Sub

Private Sub grdCommonTransactionsIndex_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1
    
End Sub

Private Sub grdCommonTransactionsIndex_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdCommonTransactionsIndex_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeyInsert Or KeyCode = vbKeyDelete Or KeyCode = vbKeySpace Then
        grdCommonTransactionsIndex.CellIcon(grdCommonTransactionsIndex.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdCommonTransactionsIndex, KeyCode, grdCommonTransactionsIndex.CurRow, "InvoiceID"))
        lblSelectedGridLines.Caption = CountSelected(grdCommonTransactionsIndex)
        If txtRefersTo.text = "1" Or txtRefersTo.text = "2" Then
            lblSelectedGridTotals.Caption = SumSelectedGridRows(grdCommonTransactionsIndex, False, "InvoiceRestAmount", "InvoiceVATAmount", "InvoiceGrossAmount")
        End If
        If txtRefersTo.text = "3" Or txtRefersTo.text = "4" Then
            lblSelectedGridTotals.Caption = SumSelectedGridRows(grdCommonTransactionsIndex, False, "InvoiceGrossAmount")
        End If
        If txtRefersTo.text = "5" Then
            lblSelectedGridTotals.Caption = SumSelectedGridRows(grdCommonTransactionsIndex, False, "Qty")
        End If
    End If

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdCommonTransactionsIndex" & txtRefersTo.text, grdCommonTransactionsIndex.LayoutCol

End Sub

Private Sub txtCategoryShortDescription_Change()

    If txtCategoryShortDescription.text = "" Then ClearFields txtCategoryID, lblCategoryDescription, txtItemID, txtItemDescription

End Sub

Private Sub txtCategoryShortDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2

End Sub

Private Sub txtCategoryShortDescription_Validate(Cancel As Boolean)

    If txtCategoryID.text = "" And txtCategoryShortDescription.text <> "" Then cmdIndex_Click 2: If txtCategoryID.text = "" Then Cancel = True

End Sub

Private Sub txtCodeShortDescription_Change()

    If txtCodeShortDescription.text = "" Then ClearFields txtCodeID, lblCodeDescription

End Sub

Private Sub txtCodeShortDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 4

End Sub

Private Sub txtCodeShortDescription_Validate(Cancel As Boolean)

    If txtCodeID.text = "" And txtCodeShortDescription.text <> "" Then cmdIndex_Click 4: If txtCodeID.text = "" Then Cancel = True

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

Private Sub txtItemDescription_Change()

    If txtItemDescription.text = "" Then ClearFields txtItemID

End Sub

Private Sub txtItemDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 3

End Sub

Private Sub txtItemDescription_Validate(Cancel As Boolean)

    If txtItemID.text = "" And txtItemDescription.text <> "" Then cmdIndex_Click 3: If txtItemID.text = "" Then Cancel = True

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

