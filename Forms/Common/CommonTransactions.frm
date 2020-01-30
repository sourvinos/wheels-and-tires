VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form CommonTransactions 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   12105
   ClientLeft      =   15
   ClientTop       =   0
   ClientWidth     =   22080
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12105
   ScaleWidth      =   22080
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   11190
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   18315
      Begin VB.Frame frmInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Caption         =   "txtInvoiceCodeCreditID"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   9465
         Left            =   8850
         TabIndex        =   13
         Top             =   525
         Width           =   8940
         Begin VB.TextBox Text36 
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
            TabIndex        =   139
            TabStop         =   0   'False
            Text            =   "Codes.CodeIsCreditID"
            Top             =   9075
            Width           =   3540
         End
         Begin VB.TextBox Text35 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            TabIndex        =   137
            TabStop         =   0   'False
            Text            =   "Invoices.InvoiceCancelTrnID"
            Top             =   4200
            Width           =   3540
         End
         Begin VB.TextBox txtInvoiceCancelTrnID 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            TabIndex        =   136
            TabStop         =   0   'False
            Text            =   "32"
            Top             =   4200
            Width           =   780
         End
         Begin VB.TextBox txtCodeDTableCredit 
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
            TabIndex        =   135
            TabStop         =   0   'False
            Text            =   "22"
            Top             =   8700
            Width           =   780
         End
         Begin VB.TextBox Text32 
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
            TabIndex        =   134
            TabStop         =   0   'False
            Text            =   "Codes.CodeDTableCredit"
            Top             =   8700
            Width           =   3540
         End
         Begin VB.TextBox Text34 
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
            TabIndex        =   130
            TabStop         =   0   'False
            Text            =   "Codes.CodeDTable"
            Top             =   8325
            Width           =   3540
         End
         Begin VB.TextBox txtCodeDTable 
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
            TabIndex        =   129
            TabStop         =   0   'False
            Text            =   "21"
            Top             =   8325
            Width           =   780
         End
         Begin VB.TextBox txtInvoiceCodeCreditID 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            TabIndex        =   128
            TabStop         =   0   'False
            Text            =   "32"
            Top             =   3825
            Width           =   780
         End
         Begin VB.TextBox Text29 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            TabIndex        =   127
            TabStop         =   0   'False
            Text            =   "Invoices.InvoiceCodeCreditID"
            Top             =   3825
            Width           =   3540
         End
         Begin VB.TextBox txtCodeIsCreditID 
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
            TabIndex        =   126
            TabStop         =   0   'False
            Text            =   "20"
            Top             =   9075
            Width           =   780
         End
         Begin VB.TextBox Text30 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            TabIndex        =   120
            TabStop         =   0   'False
            Text            =   "Invoices.InvoicePrintExtraRemarks"
            Top             =   3450
            Width           =   3540
         End
         Begin VB.TextBox txtInvoicePrintExtraRemarksID 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            TabIndex        =   119
            TabStop         =   0   'False
            Text            =   "9"
            Top             =   3450
            Width           =   780
         End
         Begin VB.TextBox txtCodeDateCheckID 
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
            TabIndex        =   118
            TabStop         =   0   'False
            Text            =   "19"
            Top             =   7950
            Width           =   780
         End
         Begin VB.TextBox Text27 
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
            TabIndex        =   117
            TabStop         =   0   'False
            Text            =   "Codes.CodeDateCheckID"
            Top             =   7950
            Width           =   3540
         End
         Begin VB.TextBox Text28 
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
            TabIndex        =   116
            TabStop         =   0   'False
            Text            =   "Codes.CodeDetailLines"
            Top             =   7575
            Width           =   3540
         End
         Begin VB.TextBox mskCodeDetailLines 
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
            TabIndex        =   115
            TabStop         =   0   'False
            Text            =   "18"
            Top             =   7575
            Width           =   780
         End
         Begin VB.TextBox txtCodePrinterID 
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
            TabIndex        =   114
            TabStop         =   0   'False
            Text            =   "17"
            Top             =   7200
            Width           =   780
         End
         Begin VB.TextBox Text12 
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
            TabIndex        =   113
            TabStop         =   0   'False
            Text            =   "Codes.CodePrinterID"
            Top             =   7200
            Width           =   3540
         End
         Begin VB.TextBox Text21 
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
            TabIndex        =   112
            TabStop         =   0   'False
            Text            =   "Codes.CodeLastDate"
            Top             =   6825
            Width           =   3540
         End
         Begin VB.TextBox mskCodeLastDate 
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
            TabIndex        =   111
            TabStop         =   0   'False
            Text            =   "16"
            Top             =   6825
            Width           =   780
         End
         Begin VB.TextBox Text25 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
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
            Left            =   4500
            TabIndex        =   110
            TabStop         =   0   'False
            Text            =   "Person.City"
            Top             =   1575
            Width           =   3540
         End
         Begin VB.TextBox txtCity 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
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
            Left            =   8100
            TabIndex        =   109
            TabStop         =   0   'False
            Text            =   "27"
            Top             =   1575
            Width           =   780
         End
         Begin VB.TextBox Text23 
            Appearance      =   0  'Flat
            BackColor       =   &H00004080&
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
            Left            =   4500
            TabIndex        =   108
            TabStop         =   0   'False
            Text            =   "TaxOffices.TaxOfficeDescription"
            Top             =   3075
            Width           =   3540
         End
         Begin VB.TextBox txtTaxOfficeDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H00004080&
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
            Left            =   8100
            TabIndex        =   107
            TabStop         =   0   'False
            Text            =   "31"
            Top             =   3075
            Width           =   780
         End
         Begin VB.TextBox txtPhones 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
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
            Left            =   8100
            TabIndex        =   106
            TabStop         =   0   'False
            Text            =   "29"
            Top             =   2325
            Width           =   780
         End
         Begin VB.TextBox Text26 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
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
            Left            =   4500
            TabIndex        =   105
            TabStop         =   0   'False
            Text            =   "Person.Phones"
            Top             =   2325
            Width           =   3540
         End
         Begin VB.TextBox txtAddress 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
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
            Left            =   8100
            TabIndex        =   104
            TabStop         =   0   'False
            Text            =   "26"
            Top             =   1200
            Width           =   780
         End
         Begin VB.TextBox Text24 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
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
            Left            =   4500
            TabIndex        =   103
            TabStop         =   0   'False
            Text            =   "Person.Address"
            Top             =   1200
            Width           =   3540
         End
         Begin VB.TextBox txtTaxNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
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
            Left            =   8100
            TabIndex        =   102
            TabStop         =   0   'False
            Text            =   "28"
            Top             =   1950
            Width           =   780
         End
         Begin VB.TextBox Text22 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
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
            Left            =   4500
            TabIndex        =   101
            TabStop         =   0   'False
            Text            =   "Person.TaxNo"
            Top             =   1950
            Width           =   3540
         End
         Begin VB.TextBox txtVATStateID 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
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
            Left            =   8100
            TabIndex        =   100
            TabStop         =   0   'False
            Text            =   "30"
            Top             =   2700
            Width           =   780
         End
         Begin VB.TextBox Text17 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
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
            Left            =   4500
            TabIndex        =   99
            TabStop         =   0   'False
            Text            =   "Person.VATStateID"
            Top             =   2700
            Width           =   3540
         End
         Begin VB.TextBox txtInvoiceInTime 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            Text            =   "7"
            Top             =   2325
            Width           =   780
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            Text            =   "Invoices.InvoiceInTime"
            Top             =   2325
            Width           =   3540
         End
         Begin VB.TextBox txtInvoiceInDate 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            TabIndex        =   47
            TabStop         =   0   'False
            Text            =   "6"
            Top             =   1950
            Width           =   780
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            TabIndex        =   46
            TabStop         =   0   'False
            Text            =   "Invoices.InvoiceInDate"
            Top             =   1950
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
            Left            =   8100
            TabIndex        =   45
            TabStop         =   0   'False
            Text            =   "23"
            Top             =   75
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
            Left            =   4500
            TabIndex        =   44
            TabStop         =   0   'False
            Text            =   "Table"
            Top             =   75
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
            Left            =   4500
            TabIndex        =   43
            TabStop         =   0   'False
            Text            =   "RefersTo"
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
            Left            =   8100
            TabIndex        =   42
            TabStop         =   0   'False
            Text            =   "24"
            Top             =   450
            Width           =   780
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            Text            =   "Invoices.InvoiceCodeID"
            Top             =   825
            Width           =   3540
         End
         Begin VB.TextBox txtInvoiceCodeID 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            TabIndex        =   40
            TabStop         =   0   'False
            Text            =   "3"
            Top             =   825
            Width           =   780
         End
         Begin VB.TextBox txtInvoiceTrnID 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            TabIndex        =   38
            TabStop         =   0   'False
            Text            =   "Invoices.InvoiceTrnID"
            Top             =   75
            Width           =   3540
         End
         Begin VB.TextBox txtInvoicePersonID 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            TabIndex        =   37
            TabStop         =   0   'False
            Text            =   "2"
            Top             =   450
            Width           =   780
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            Text            =   "Invoices.InvoicePersonID"
            Top             =   450
            Width           =   3540
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            Text            =   "Invoices.InvoiceDeliveryPointID"
            Top             =   1200
            Width           =   3540
         End
         Begin VB.TextBox txtInvoiceDeliveryPointID 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            Text            =   "4"
            Top             =   1200
            Width           =   780
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            Text            =   "Invoices.InvoicePaymentWayID"
            Top             =   1575
            Width           =   3540
         End
         Begin VB.TextBox txtInvoicePaymentWayID 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            TabIndex        =   32
            TabStop         =   0   'False
            Text            =   "5"
            Top             =   1575
            Width           =   780
         End
         Begin VB.TextBox txtInvoiceIsInvoiced 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            TabIndex        =   31
            TabStop         =   0   'False
            Text            =   "8"
            Top             =   2700
            Width           =   780
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            Text            =   "Invoices.InvoiceIsInvoiced"
            Top             =   2700
            Width           =   3540
         End
         Begin VB.TextBox txtInvoiceIsPrinted 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            TabIndex        =   29
            TabStop         =   0   'False
            Text            =   "9"
            Top             =   3075
            Width           =   780
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            TabIndex        =   28
            TabStop         =   0   'False
            Text            =   "Invoices.InvoiceIsPrinted"
            Top             =   3075
            Width           =   3540
         End
         Begin VB.TextBox Text9 
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
            TabIndex        =   27
            TabStop         =   0   'False
            Text            =   "Codes.CodeDetailsID"
            Top             =   4575
            Width           =   3540
         End
         Begin VB.TextBox txtCodeDetailsID 
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
            TabIndex        =   26
            TabStop         =   0   'False
            Text            =   "10"
            Top             =   4575
            Width           =   780
         End
         Begin VB.TextBox Text13 
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
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   "Codes.CodeHandID"
            Top             =   4950
            Width           =   3540
         End
         Begin VB.TextBox txtCodeHandID 
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
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   "11"
            Top             =   4950
            Width           =   780
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
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
            Left            =   4500
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   "Person.Profession"
            Top             =   825
            Width           =   3540
         End
         Begin VB.TextBox txtProfession 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
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
            Left            =   8100
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   "25"
            Top             =   825
            Width           =   780
         End
         Begin VB.TextBox txtCodeLastNo 
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
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "12"
            Top             =   5325
            Width           =   780
         End
         Begin VB.TextBox Text16 
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
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "Codes.CodeLastNo"
            Top             =   5325
            Width           =   3540
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
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "Codes.CodeInventoryQty"
            Top             =   5700
            Width           =   3540
         End
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
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "13"
            Top             =   5700
            Width           =   780
         End
         Begin VB.TextBox Text18 
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
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "Codes.CodeInventoryValue"
            Top             =   6075
            Width           =   3540
         End
         Begin VB.TextBox txtCodeInventoryValue 
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
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   "14"
            Top             =   6075
            Width           =   780
         End
         Begin VB.TextBox txtCodeTransformID 
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
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "15"
            Top             =   6450
            Width           =   780
         End
         Begin VB.TextBox Text19 
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
            TabIndex        =   14
            TabStop         =   0   'False
            Text            =   "Codes.CodeTransformID"
            Top             =   6450
            Width           =   3540
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   8325
            Top             =   3450
            _ExtentX        =   953
            _ExtentY        =   953
            IconSizeX       =   26
            IconSizeY       =   32
            Size            =   14064
            Images          =   "CommonTransactions.frx":0000
            Version         =   131072
            KeyCount        =   4
            Keys            =   "ÿÿÿ"
         End
      End
      Begin VB.Frame frmDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   450
         TabIndex        =   90
         Top             =   7575
         Width           =   15840
         Begin UserControls.newText txtInvoiceDestinationSite 
            Height          =   465
            Left            =   11775
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   300
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   820
            Enabled         =   0   'False
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
         Begin UserControls.newText txtInvoiceTransportReason 
            Height          =   465
            Left            =   300
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   300
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   820
            Enabled         =   0   'False
            ForeColor       =   0
            MaxLength       =   40
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
         Begin UserControls.newText txtInvoiceTransportWay 
            Height          =   465
            Left            =   7950
            TabIndex        =   93
            TabStop         =   0   'False
            Top             =   300
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   820
            Enabled         =   0   'False
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
         Begin UserControls.newText txtInvoiceLoadingSite 
            Height          =   465
            Left            =   4125
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   300
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   820
            Enabled         =   0   'False
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
         Begin VB.Label lblSimple 
            BackColor       =   &H000000C0&
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
            Index           =   11
            Left            =   7950
            TabIndex        =   98
            Top             =   0
            Width           =   1365
         End
         Begin VB.Label lblSimple 
            BackColor       =   &H000000C0&
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
            Index           =   10
            Left            =   4125
            TabIndex        =   97
            Top             =   0
            Width           =   1365
         End
         Begin VB.Label lblSimple 
            BackColor       =   &H000000C0&
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
            Index           =   9
            Left            =   11775
            TabIndex        =   96
            Top             =   0
            Width           =   1365
         End
         Begin VB.Label lblSimple 
            BackColor       =   &H000000C0&
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
            Index           =   12
            Left            =   300
            TabIndex        =   95
            Top             =   0
            Width           =   1365
         End
      End
      Begin VB.Frame frmTotals 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1590
         Left            =   450
         TabIndex        =   73
         Top             =   8400
         Width           =   9390
         Begin VB.TextBox eLine 
            Alignment       =   2  'Center
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
            Left            =   225
            TabIndex        =   138
            TabStop         =   0   'False
            Text            =   "eLine"
            Top             =   1200
            Width           =   8940
         End
         Begin UserControls.newInteger mskTotalQty 
            Height          =   540
            Left            =   225
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   600
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   953
            Enabled         =   0   'False
            Alignment       =   1
            ForeColor       =   0
            Text            =   "9.999"
            BackColor       =   0
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
         Begin UserControls.newFloat mskTotalPreDiscount 
            Height          =   540
            Left            =   1350
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   600
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   953
            Enabled         =   0   'False
            Alignment       =   1
            ForeColor       =   0
            Text            =   "99.999,99"
            BackColor       =   0
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
         Begin UserControls.newFloat mskDiscount 
            Height          =   540
            Left            =   2475
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   600
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   953
            Enabled         =   0   'False
            Alignment       =   1
            ForeColor       =   0
            Text            =   "99.999,99"
            BackColor       =   0
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
         Begin UserControls.newFloat mskTransDiscount 
            Height          =   540
            Left            =   3600
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   600
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   953
            Enabled         =   0   'False
            Alignment       =   1
            ForeColor       =   0
            Text            =   "99.999,99"
            BackColor       =   0
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
         Begin UserControls.newFloat mskTotalRestAmount 
            Height          =   540
            Left            =   4725
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   600
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   953
            Enabled         =   0   'False
            Alignment       =   1
            ForeColor       =   0
            Text            =   "99.999,99"
            BackColor       =   0
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
         Begin UserControls.newFloat mskExtraCharges 
            Height          =   540
            Left            =   5850
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   600
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   953
            Enabled         =   0   'False
            Alignment       =   1
            ForeColor       =   0
            Text            =   "99.999,99"
            BackColor       =   0
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
         Begin UserControls.newFloat mskTotalVAT 
            Height          =   540
            Left            =   6975
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   600
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   953
            Enabled         =   0   'False
            Alignment       =   1
            ForeColor       =   0
            Text            =   "99.999,99"
            BackColor       =   0
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
         Begin UserControls.newFloat mskTotalGross 
            Height          =   540
            Left            =   8100
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   600
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   953
            Enabled         =   0   'False
            Alignment       =   1
            ForeColor       =   8388736
            Text            =   "19.999,99"
            BackColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Ubuntu Condensed"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblSimple 
            Alignment       =   2  'Center
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
            Index           =   19
            Left            =   8100
            TabIndex        =   88
            Top             =   225
            Width           =   1065
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblSimple 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            Caption         =   "ΦΠΑ"
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
            Left            =   6975
            TabIndex        =   87
            Top             =   225
            Width           =   1065
         End
         Begin VB.Label lblSimple 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            Caption         =   "Λοιπές χρεώσεις"
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
            Height          =   540
            Index           =   17
            Left            =   5850
            TabIndex        =   86
            Top             =   0
            Width           =   1065
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblSimple 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            Caption         =   "Υπόλοιπο αξίας"
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
            Height          =   540
            Index           =   16
            Left            =   4725
            TabIndex        =   85
            Top             =   0
            Width           =   1065
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblSimple 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            Caption         =   "Επιπλέον ποσό έκπτωσης"
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
            Height          =   540
            Index           =   15
            Left            =   3600
            TabIndex        =   84
            Top             =   0
            Width           =   1065
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblSimple 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            Caption         =   "Ποσό έκπτωσης"
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
            Height          =   540
            Index           =   14
            Left            =   2475
            TabIndex        =   83
            Top             =   0
            Width           =   1065
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblSimple 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            Caption         =   "Αξία"
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
            Left            =   1350
            TabIndex        =   82
            Top             =   225
            Width           =   1065
         End
         Begin VB.Label lblSimple 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            Caption         =   "Ποσότητα"
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
            Left            =   225
            TabIndex        =   81
            Top             =   225
            Width           =   1065
         End
      End
      Begin VB.Frame frmButtonFrame 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   450
         TabIndex        =   50
         Top             =   10425
         Width           =   8940
         Begin GurhanButtonOCX.GurhanButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   51
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
            TabIndex        =   52
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
            Index           =   5
            Left            =   7350
            TabIndex        =   53
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
            TabIndex        =   54
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
            TabIndex        =   55
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
            TabIndex        =   56
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
      Begin iGrid300_10Tec.iGrid grdCommonTransactions 
         Height          =   4215
         Left            =   450
         TabIndex        =   12
         Top             =   3075
         Width           =   17415
         _ExtentX        =   30718
         _ExtentY        =   7435
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
      Begin UserControls.newDate mskInvoiceIssueDate 
         Height          =   465
         Left            =   2175
         TabIndex        =   1
         Top             =   900
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   820
         Enabled         =   0   'False
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
         Left            =   2175
         TabIndex        =   3
         Top             =   1950
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   820
         Enabled         =   0   'False
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
         Left            =   7950
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1425
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
         PicNormal       =   "CommonTransactions.frx":3710
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   1
         Left            =   8400
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1425
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
         PicNormal       =   "CommonTransactions.frx":3CAA
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newText txtPersonDescription 
         Height          =   465
         Left            =   2175
         TabIndex        =   2
         Top             =   1425
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   820
         Enabled         =   0   'False
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
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   2
         Left            =   3750
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1950
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
         PicNormal       =   "CommonTransactions.frx":4244
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   3
         Left            =   4200
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1950
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
         PicNormal       =   "CommonTransactions.frx":47DE
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newText txtDeliveryPointDescription 
         Height          =   465
         Left            =   11475
         TabIndex        =   9
         Top             =   1425
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   820
         Enabled         =   0   'False
         ForeColor       =   0
         MaxLength       =   40
         Text            =   "τοπος παραλαβης"
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
      Begin UserControls.newText txtPaymentWayDescription 
         Height          =   465
         Left            =   11475
         TabIndex        =   10
         Top             =   1950
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   820
         Enabled         =   0   'False
         ForeColor       =   0
         MaxLength       =   40
         Text            =   "τροπος πληρωμης"
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
      Begin UserControls.newText txtInvoicePlates 
         Height          =   465
         Left            =   11475
         TabIndex        =   7
         Top             =   900
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   820
         Enabled         =   0   'False
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
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   4
         Left            =   17025
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   900
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
         PicNormal       =   "CommonTransactions.frx":4D78
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   7
         Left            =   17025
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   1950
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
         PicNormal       =   "CommonTransactions.frx":5312
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   8
         Left            =   17475
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   1950
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
         PicNormal       =   "CommonTransactions.frx":58AC
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newText txtInvoiceNo 
         Height          =   465
         Left            =   2175
         TabIndex        =   4
         Top             =   2475
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
      Begin UserControls.newText txtInvoicePrintExtraRemarks 
         Height          =   465
         Left            =   16350
         TabIndex        =   8
         Top             =   900
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
         Left            =   17025
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   1425
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
         PicNormal       =   "CommonTransactions.frx":5E46
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   6
         Left            =   17475
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   1425
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
         PicNormal       =   "CommonTransactions.frx":63E0
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newText txtInvoiceRemarks 
         Height          =   465
         Left            =   11475
         TabIndex        =   11
         Top             =   2475
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   820
         Enabled         =   0   'False
         ForeColor       =   0
         MaxLength       =   100
         Text            =   "παρατηρησεις"
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
         Index           =   9
         Left            =   8850
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   1425
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   820
         BackColor       =   16777215
         ButtonShape     =   3
         ButtonStyle     =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210688
         PicNormal       =   "CommonTransactions.frx":697A
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newText txtCodeShortDescriptionCredit 
         Height          =   465
         Left            =   5850
         TabIndex        =   5
         Top             =   2475
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   820
         Enabled         =   0   'False
         Alignment       =   2
         ForeColor       =   0
         MaxLength       =   4
         Text            =   "ΑΑΑΑ"
         BackColor       =   4
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
      Begin UserControls.newText txtInvoiceNoCredit 
         Height          =   465
         Left            =   7425
         TabIndex        =   6
         Top             =   2475
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   820
         Alignment       =   2
         ForeColor       =   0
         MaxLength       =   7
         Text            =   "9999"
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
         Index           =   10
         Left            =   6600
         TabIndex        =   131
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
         PicNormal       =   "CommonTransactions.frx":6F14
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin VB.Label lblSimple 
         Alignment       =   2  'Center
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
         Index           =   22
         Left            =   7050
         TabIndex        =   133
         Top             =   2550
         Width           =   315
      End
      Begin VB.Label lblSimple 
         BackColor       =   &H000080FF&
         Caption         =   "Παραστατικό που ακυρώνει"
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
         Left            =   3750
         TabIndex        =   132
         Top             =   2550
         Width           =   2040
      End
      Begin VB.Label lblSimple 
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
         Left            =   9675
         TabIndex        =   124
         Top             =   2550
         Width           =   1365
      End
      Begin VB.Label lblSimple 
         BackColor       =   &H000080FF&
         Caption         =   "Τρόπος πληρωμής"
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
         Left            =   9675
         TabIndex        =   70
         Top             =   2025
         Width           =   1365
      End
      Begin VB.Label lblSimple 
         BackColor       =   &H000080FF&
         Caption         =   "Εκτύπωση ΥγΣΜ"
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
         Left            =   15075
         TabIndex        =   121
         Top             =   975
         Width           =   1215
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   915
         Index           =   5
         Left            =   15975
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblCodeDescription 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ΔΕΛΤΙΟ ΑΠΟΣΤΟΛΗΣ - ΤΙΜΟΛΟΓΙΟ ΠΩΛΗΣΗΣ"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   390
         Left            =   6600
         TabIndex        =   72
         Top             =   300
         Width           =   11265
      End
      Begin VB.Label lblSimple 
         BackColor       =   &H000080FF&
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
         Index           =   8
         Left            =   9675
         TabIndex        =   71
         Top             =   1500
         Width           =   1365
      End
      Begin VB.Label lblSimple 
         BackColor       =   &H000080FF&
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
         Index           =   6
         Left            =   9675
         TabIndex        =   69
         Top             =   975
         Width           =   1365
      End
      Begin VB.Label lblSimple 
         BackColor       =   &H000080FF&
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
         TabIndex        =   68
         Top             =   1500
         Width           =   1290
      End
      Begin VB.Label lblSimple 
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
         TabIndex        =   67
         Top             =   2025
         Width           =   1290
      End
      Begin VB.Label lblSimple 
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
         TabIndex        =   66
         Top             =   2550
         Width           =   1290
      End
      Begin VB.Label lblSimple 
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
         TabIndex        =   65
         Top             =   975
         Width           =   1290
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   2865
         Index           =   8
         Left            =   17850
         Top             =   1275
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Κινήσεις"
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
         Left            =   450
         TabIndex        =   64
         Top             =   0
         Width           =   2040
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   915
         Index           =   2
         Left            =   2700
         Top             =   0
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
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   465
         Left            =   8025
         Top             =   9975
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   1140
         Index           =   4
         Left            =   9225
         Top             =   1500
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   1140
         Index           =   1
         Left            =   1725
         Top             =   1200
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   1665
         Index           =   3
         Left            =   11025
         Top             =   1200
         Visible         =   0   'False
         Width           =   465
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
Attribute VB_Name = "CommonTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim blnStatus As Boolean
Dim blnError As Boolean
Dim blnGridEditInProgress As Boolean
Dim lngTrnID As Long
Dim aInvoicesArray() As String



Private Function AddGridLines()

    Dim lngRowsToAdd As Long
    
    lngRowsToAdd = Val(mskCodeDetailLines.text)
    
    If lngRowsToAdd - grdCommonTransactions.RowCount <> 0 Then
        grdCommonTransactions.AddRow , , , , , , , lngRowsToAdd - grdCommonTransactions.RowCount
    End If

End Function

Private Function CheckForValidSalesInvoiceNo()

    'Local μεταβλητές
    Dim intIndex As Byte
    Dim strThisQuery As String
    Dim strParameters As String
    Dim strParFields As String
    Dim strThisParameter As String
    Dim strOrder As String
    Dim strLogic As String
    Dim arrQuery() As Variant
    Dim strSQL As String
    Dim lngRow As Long
    Dim rstInvoices As Recordset
    Dim intYear As Integer
    Dim strInvoiceNo As String
    Dim lngCodeID As Long
    
    'Αρχικές τιμές
    intIndex = 0
    lngRow = 0
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Κύριο SQL - 1ο πέρασμα - Ελέγχω αν ο αριθμός που έχει δοθεί από το Codes.CodeLastNo + 1 είναι ήδη καταχωρημένος
    strSQL = "SELECT InvoiceIssueDate, InvoiceCodeID, InvoiceNo " _
        & "FROM Invoices "
    
    'Χρήση
    strThisParameter = "intYear Integer"
    strThisQuery = "Year(InvoiceIssueDate) = intYear"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Year(mskInvoiceIssueDate.text)
    
    'ID στοιχείου
    strThisParameter = "lngInvoiceID Long"
    strThisQuery = "InvoiceCodeID = lngInvoiceID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtInvoiceCodeID.text)
    
    'Νο στοιχείου
    strThisParameter = "lngInvoiceNo Long"
    strThisQuery = "InvoiceNo = lngInvoiceNo"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtInvoiceNo.text)
    
    'Προσθέτω τα κριτήρια
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
    End If
    
    TempQuery.SQL = strSQL
    
    For intIndex = 1 To UBound(arrQuery)
        TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
    Next intIndex
    
    'Ανοίγω το recordset
    Set rstInvoices = TempQuery.OpenRecordset()
    
    'Αν βρήκα εγγραφές, υπάρχει λάθος
    If rstInvoices.RecordCount > 0 Then
        rstInvoices.MoveLast
        CheckForValidSalesInvoiceNo = True
        Exit Function
    End If
    
    'Κύριο SQL - 2ο πέρασμα - Ελέγχω αν ο αριθμός που έχει δοθεί από το Codes.CodeLastNo + 1 είναι ο επόμενος από τον τελευταίο καταχωρημένο στο Invoices.InvoiceNo
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    strSQL = "SELECT InvoiceIssueDate, InvoiceCodeID, InvoiceNo " _
        & "FROM Invoices "
    
    strOrder = " ORDER BY InvoiceIssueDate, InvoiceNo"
    
    intIndex = 0
    strParameters = ""
    strParFields = ""
    
    strThisParameter = "intYear Integer"
    strThisQuery = "Year(InvoiceIssueDate) = intYear"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Year(mskInvoiceIssueDate.text)
    
    'ID στοιχείου
    strThisParameter = "lngInvoiceID Long"
    strThisQuery = "InvoiceCodeID = lngInvoiceID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtInvoiceCodeID.text)
    
    'Προσθέτω τα κριτήρια
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
    End If
    
    TempQuery.SQL = strSQL & strOrder
    
    For intIndex = 1 To UBound(arrQuery)
        TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
    Next intIndex
    
    'Ανοίγω το recordset
    Set rstInvoices = TempQuery.OpenRecordset()
    
    'Ελέγχω για διπλοεγγραφές
    With rstInvoices
        If .RecordCount > 0 Then
            .MoveLast
            If rstInvoices!InvoiceNo + 1 <> txtInvoiceNo.text Then
                CheckForValidSalesInvoiceNo = True
            End If
        Else
            CheckForValidSalesInvoiceNo = False
        End If
        .Close
    End With
    
    Exit Function

UpdateSQLString:
    intIndex = intIndex + 1
    strParameters = IIf(intIndex > 1, strParameters & ", ", strParameters)
    strParFields = IIf(intIndex > 1, strParFields & strLogic, strParFields)
    strParameters = strParameters & strThisParameter
    strParFields = strParFields & strThisQuery
    ReDim Preserve arrQuery(intIndex)
    Return

End Function

Function ColorizeRowsWhenItemIsNotGiven(myRow As Long)

    Dim lngRow As Long
    Dim lngCol As Long
    
    grdCommonTransactions.Redraw = False
    
    If myRow <> 0 Then
        If grdCommonTransactions.CellText(myRow, "CategoryID") = "" And grdCommonTransactions.CellText(myRow, "ItemID") = "" Then
            For lngCol = 5 To grdCommonTransactions.colCount
                grdCommonTransactions.CellForeColor(myRow, lngCol) = vbBlack
            Next lngCol
        Else
            For lngCol = 5 To grdCommonTransactions.colCount
                grdCommonTransactions.CellForeColor(myRow, lngCol) = vbWhite
            Next lngCol
        End If
        grdCommonTransactions.Redraw = True
        Exit Function
    End If
    
    For lngRow = 1 To grdCommonTransactions.RowCount
        If grdCommonTransactions.CellText(lngRow, "CategoryID") = "" And grdCommonTransactions.CellText(lngRow, "ItemID") = "" Then
            For lngCol = 5 To grdCommonTransactions.colCount
                grdCommonTransactions.CellForeColor(lngRow, lngCol) = vbBlack
            Next lngCol
        Else
            For lngCol = 5 To grdCommonTransactions.colCount
                grdCommonTransactions.CellForeColor(lngRow, lngCol) = vbWhite
            Next lngCol
        End If
    Next lngRow
    
    grdCommonTransactions.Redraw = True

End Function

Private Function CreateELine()

    'Συμβολοσειρά αναγνώρισης αρχής
    'ΑΦΜ δικο μας
    'ΑΦΜ πελάτη
    'Νο κάρτας αποδείξεων
    'Τύπος παραστατικού # Τύπος παραστατικού που ακυρώνεται # Νο παραστατικού που ακυρώνεται
    'Σειρά παραστατικού
    'Αύξουδα αρίθμηση παραστατικού
    'Καθαρό ποσό συντελεστή Α
    'Καθαρό ποσό συντελεστή Β
    'Καθαρό ποσό συντελεστή Γ
    'Καθαρό ποσό συντελεστή Δ
    'Καθαρό ποσό συντελεστή Ε
    'ΦΠΑ συντελεστή Α
    'ΦΠΑ συντελεστή Β
    'ΦΠΑ συντελεστή Γ
    'ΦΠΑ συντελεστή Δ
    'ΦΠΑ συντελεστή Ε
    'Σύνολο αξίας
    'Κωδικός νομίσματος
    'Συμβολοσειρά αναγνώρισης τέλους
    
    eLine.text = _
        "$$$" + _
        "099863549" + ";" + _
        txtTaxNo.text + ";" + _
        "0" + ";" + _
         txtCodeDTable.text + ";" + _
         IIf(txtCodeIsCreditID.text = "1", "#" + txtCodeDTableCredit.text + "#" + txtInvoiceNoCredit.text, "") + _
        "" + ";" + _
        txtInvoiceNo.text + ";" + _
        "0" + ";" + _
        "0" + ";" + _
        TransformAmount(mskTotalRestAmount.text) + ";" + _
        "0" + ";" + _
        "0" + ";" + _
        "0" + ";" + _
        "0" + ";" + _
        TransformAmount(mskTotalVAT.text) + ";" + _
        "0" + ";" + _
        "0" + ";" + _
        TransformAmount(mskTotalGross.text) + ";" + _
        "0" + _
        "$$$"
        
        CreateELine = eLine.text

End Function

Function DoSharedStuff(myInvoiceTrnID, myWindowTitle, myTable, myRefersTo)

    FillCellWithSomething grdCommonTransactions, 0, 0, "5,6,7,8,9,10,12,13,14,15,16"
    FindInvoicesWithTrnID myInvoiceTrnID, myWindowTitle, myTable, myRefersTo
    FindItemsWithTrnID myInvoiceTrnID
    AddGridLines
    
    If txtCodeHandID.text = "1" Or txtCodeDateCheckID.text = "0" Then
        
        EnableFields mskInvoiceIssueDate, txtPersonDescription, txtCodeShortDescription, txtInvoiceNo, txtInvoicePrintExtraRemarks, txtDeliveryPointDescription, txtPaymentWayDescription, txtInvoicePlates, txtInvoiceRemarks, txtInvoiceTransportReason, txtInvoiceTransportWay, txtInvoiceLoadingSite, txtInvoiceDestinationSite, txtCodeShortDescriptionCredit, txtInvoiceNoCredit
        EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8), cmdIndex(9)
        EnableFields mskDiscount, mskTransDiscount, mskTotalRestAmount, mskExtraCharges, mskTotalVAT
        
        UpdateButtons Me, 5, 0, 1, 1, 0, 1, 0
        
        EnableOrDisableSpecialFields
            
    End If
    
    'Αν το παραστατικό είναι μηχανογραφικό πωλήσεων - η ημερομηνία έκδοσης είναι διαφορετική της σημερινής - γίνεται έλεγχος ημερομηνίας
    If txtCodeHandID.text = "0" And txtRefersTo.text = "2" And CDate(mskInvoiceIssueDate.text) <> Date And txtCodeDateCheckID.text = "1" Then
        'Τα πάντα απαγορεύονται
        UpdateButtons Me, 5, 0, 0, 0, 0, 0, 1
    End If
    
    'Αν το παραστατικό είναι μηχανογραφικό πωλήσεων - η ημερομηνία έκδοσης είναι ίση με τη σημερινή
    If txtCodeHandID.text = "0" And txtRefersTo.text = "2" And CDate(mskInvoiceIssueDate.text) = Date Then
        'Επιτρέπεται η επανεκτύπωση
        UpdateButtons Me, 5, 0, 1, 0, 0, 0, 1
        cmdButton(1).Caption = "Επανεκτύπωση"
    End If
    
    ColorizeRowsWhenItemIsNotGiven 0
    Tag = "False"
    ShowOrHideDetailsFrame

End Function

Function EnableOrDisableSpecialFields()

    txtInvoicePlates.Enabled = IIf(txtRefersTo.text = "1", False, True) 'Αρ. κυκλοφορίας
    txtInvoicePrintExtraRemarks.Enabled = IIf(txtRefersTo.text = "1", False, True) 'ΥγΣΜ
    txtInvoicePrintExtraRemarks.text = IIf(txtRefersTo.text = "1", "", txtInvoicePrintExtraRemarks.text) 'ΥγΣΜ
    cmdIndex(4).Enabled = IIf(txtRefersTo.text = "1", False, True) 'ΥγΣΜ
    
    txtInvoiceDeliveryPointID.text = IIf(txtRefersTo.text = "1", txtInvoiceDeliveryPointID.text, "1") 'Τόπος παραλαβής
    txtDeliveryPointDescription.Enabled = IIf(txtRefersTo.text = "1", True, False) 'Τόπος παραλαβής
    cmdIndex(5).Enabled = IIf(txtRefersTo.text = "1", True, False) 'Τόπος παραλαβής
    cmdIndex(6).Enabled = IIf(txtRefersTo.text = "1", True, False) 'Τόπος παραλαβής

End Function

Private Function FindPersonDetails(lngPersonID)

    Dim tmpRecordset As Recordset
    
    Set tmpRecordset = NewCheckForMatch("CommonDB", "ID, Description, TaxNo, Profession, Address, City, Phones, VATStateID, TaxOfficeDescription, CountryShortDescription", _
        "((" & txtTable.text, _
        "INNER JOIN TaxOffices ON " & txtTable.text & ".TaxOfficeID = TaxOffices.TaxOfficeID) " & _
        "INNER JOIN Countries ON " & txtTable.text & ".CountryID = Countries.CountryID)", "ID = " & lngPersonID, "", "ID")
    
    If tmpRecordset.RecordCount = 1 Then
        txtInvoicePersonID.text = tmpRecordset!ID
        txtPersonDescription.text = tmpRecordset!Description
        txtProfession.text = tmpRecordset!Profession
        txtAddress.text = tmpRecordset!Address
        txtCity.text = tmpRecordset!City
        txtPhones.text = tmpRecordset!Phones
        txtVATStateID.text = tmpRecordset!VATStateID
        txtTaxNo.text = tmpRecordset!CountryShortDescription & " " & tmpRecordset!TaxNo
        txtTaxOfficeDescription.text = tmpRecordset!TaxOfficeDescription
    End If
    
End Function

Private Function HideCredits()

    lblSimple(21).Visible = False
    lblSimple(22).Visible = False
    cmdIndex(10).Visible = False
    txtCodeShortDescriptionCredit.Visible = False
    txtInvoiceNoCredit.Visible = False

End Function


Private Function ItemDescriptionAndManufacturer(strItemDescription As String, strManufacturerDescription As String)

    Dim intMaxLength
    Dim intItemDescriptionLength As Integer
    Dim intManufacturerDescriptionLength As Integer
    Dim intCombinedLength As Integer
    Dim tmpRecordset As Recordset
    
    Dim strReturnString As String
    
    Set tmpRecordset = CheckForMatch("CommonDB", strManufacturerDescription, "Manufacturers", "ManufacturerDescription", "String", 1, "ManufacturerDescription")
    If tmpRecordset.RecordCount > 0 Then
        If tmpRecordset!ManufacturerIsShownID = 0 Then
            strManufacturerDescription = ""
        End If
    End If
    
    intMaxLength = 41
    
    intItemDescriptionLength = Len(strItemDescription)
    intManufacturerDescriptionLength = Len(strManufacturerDescription)
    
    intCombinedLength = intItemDescriptionLength + 1 + intManufacturerDescriptionLength
    
    If intCombinedLength > 41 Then 'Αν το συνδιασμένο μήκος είναι > 41
        If intItemDescriptionLength > 36 Then 'Αν το μήκος της περιγραφής είναι > 36
            strItemDescription = Left(strItemDescription, 36) 'Κρατάω τους πρώτους 36 χαρακτήρες της περιγραφής
            strManufacturerDescription = Left(strManufacturerDescription, 4) 'Κολάω τους 4 πρώτους χαρακτήρες του κατασκευαστή
        End If
        If intItemDescriptionLength <= 36 Then 'Αν το μήκος της περιγραφής είναι <= 36
            If intManufacturerDescriptionLength + 1 + intItemDescriptionLength > 41 Then 'Αν το μήκος του κατασκευαστή + το μήκος της περιγραφής + 1 ξεπερνάει τους 41 χαρακτήρες
                strManufacturerDescription = Left(strManufacturerDescription, 41 - 1 - intItemDescriptionLength) 'Κρατάω τους πρώτους χαρακτήρες του κατασκευαστή μέχρι να έχω συνδιασμένο μήκος = 41
            End If
        End If
        strReturnString = strItemDescription + " " + strManufacturerDescription
    End If
    
    If intCombinedLength <= 41 Then
        strReturnString = strItemDescription + " " + strManufacturerDescription
    End If
            
    ItemDescriptionAndManufacturer = Trim(strReturnString)
    
End Function

Private Function SeekInvoiceToCancel()

    Dim intIndex As Byte
    Dim strThisQuery As String
    Dim strParameters As String
    Dim strParFields As String
    Dim strThisParameter As String
    Dim strOrder As String
    Dim strLogic As String
    Dim arrQuery() As Variant
    Dim strSQL As String
    
    Dim rstRecordset As Recordset
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    strSQL = "SELECT InvoiceID FROM Invoices "
    
    'Εκδοση
    strThisParameter = "datIssue Date"
    strThisQuery = "Invoices.InvoiceIssueDate = datIssue"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = CDate(mskInvoiceIssueDate.text)
    
    'Παραστατικό
    strThisParameter = "intCodeID Integer"
    strThisQuery = "Invoices.InvoiceCodeID = intCodeID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtInvoiceCodeCreditID.text)
    
    'Νο παραστατικού
    strThisParameter = "intInvoiceNo Integer"
    strThisQuery = "Invoices.InvoiceNo = intInvoiceNo"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtInvoiceNoCredit.text)
    
    'Συναλλασόμενος
    strThisParameter = "intPersonID Integer"
    strThisQuery = "Invoices.InvoicePersonID = intPersonID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtInvoicePersonID.text)
    
    'Προσθέτω τα κριτήρια
    strParameters = "PARAMETERS " & strParameters & "; "
    strParFields = "WHERE " & strParFields
    strSQL = strParameters & strSQL & strParFields
    
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
    If rstRecordset.RecordCount = 0 Then
        If DisplayMessage(29, 4, 1, "", "") Then
            ClearFields txtCodeShortDescriptionCredit, txtInvoiceNoCredit, txtInvoiceCodeCreditID, txtCodeDTableCredit
            SeekInvoiceToCancel = False
        End If
    Else
        txtInvoiceCancelTrnID.text = rstRecordset(0)
        SeekInvoiceToCancel = True
    End If
    
    Exit Function
    
UpdateSQLString:
    intIndex = intIndex + 1
    strParameters = IIf(intIndex > 1, strParameters & ", ", strParameters)
    strParFields = IIf(intIndex > 1, strParFields & strLogic, strParFields)
    strParameters = strParameters & strThisParameter
    strParFields = strParFields & strThisQuery
    ReDim Preserve arrQuery(intIndex)
    
    Return

End Function

Private Function ShowCredits()

    lblSimple(21).Visible = True
    lblSimple(22).Visible = True
    cmdIndex(10).Visible = True
    txtCodeShortDescriptionCredit.Visible = True
    txtInvoiceNoCredit.Visible = True

End Function

Private Function TransformAmount(amount)
    
    Dim amountA As String

    amountA = Replace(amount, ".", "")
    amountA = Replace(amountA, ",", ".")
    
    TransformAmount = amountA

End Function

Private Function TransformInvoices()

    Dim intLoop As Integer
    Dim rsInvoices As Recordset
    
    Set rsInvoices = CommonDB.OpenRecordset("Invoices")
    
    If txtCodeHandID.text = "1" Or (txtCodeHandID.text = "0" And blnStatus) Then
    
        For intLoop = 1 To UBound(aInvoicesArray)
            With rsInvoices
                .Index = "TrnID"
                .Seek "=", aInvoicesArray(intLoop, 2)
                If Not .NoMatch Then
                    .Edit
                    !InvoiceIsInvoiced = 2
                    .Update
                End If
            End With
        Next intLoop
    
    End If
    
    rsInvoices.Close

End Function

Function UpdateArrayWithInvoicesToTransform()
        
    'Local μεταβλητές
    Dim intLoop As Byte
    Dim intUpper As Integer
    Dim intArrayindex As Integer
    Dim lngRow As Long
    
    'Αρχικές τιμές
    intUpper = 1
    intArrayindex = 1
    blnStatus = True
    
    With CommonPendingInvoices.grdCommonPendingInvoices
        .Sort ("Order")
        'Φτιάχνω τον πίνακα
        For lngRow = 1 To .RowCount
            If .CellIcon(lngRow, "Selected") > 0 Then
                ReDim aInvoicesArray(intUpper, 11)
                intUpper = intUpper + 1
            End If
        Next lngRow
        'Γεμίζω τον πίνακα
        For lngRow = 1 To .RowCount
            If .CellIcon(lngRow, "Selected") > 0 Then
                aInvoicesArray(intArrayindex, 1) = .CellText(lngRow, "InvoiceNo")
                aInvoicesArray(intArrayindex, 2) = .CellText(lngRow, "InvoiceTrnID")
                aInvoicesArray(intArrayindex, 3) = .CellText(lngRow, "PersonID")
                aInvoicesArray(intArrayindex, 4) = .CellText(lngRow, "PersonDescription")
                aInvoicesArray(intArrayindex, 5) = .CellText(lngRow, "InvoiceIssueDate")
                aInvoicesArray(intArrayindex, 6) = .CellText(lngRow, "DeliveryPointID")
                aInvoicesArray(intArrayindex, 7) = .CellText(lngRow, "DeliveryPointDescription")
                aInvoicesArray(intArrayindex, 8) = .CellText(lngRow, "PaymentWayID")
                aInvoicesArray(intArrayindex, 9) = .CellText(lngRow, "PaymentWayDescription")
                aInvoicesArray(intArrayindex, 10) = .CellText(lngRow, "InvoiceIssueDate")
                aInvoicesArray(intArrayindex, 11) = .CellText(lngRow, "InvoiceRemarks")
                intArrayindex = intArrayindex + 1
            End If
        Next lngRow
    End With

End Function

Function UpdateRemarksFieldWithInvoices()

    Dim intLoop As Byte
    Dim strInvoiceRemarks As String
    
    strInvoiceRemarks = "ΔΑ "
    
    For intLoop = 1 To UBound(aInvoicesArray)
        strInvoiceRemarks = strInvoiceRemarks & aInvoicesArray(intLoop, 1) & ", "
    Next intLoop
    
    txtInvoiceRemarks.text = Left(strInvoiceRemarks, Len(strInvoiceRemarks) - 2)

End Function

Function UpdateGridWithItems()

    Dim intLoop As Integer
    
    For intLoop = 1 To UBound(aInvoicesArray)
        FindItemsWithTrnID Val(aInvoicesArray(intLoop, 2))
    Next intLoop
    
End Function

Function UpdateHeaders()

    mskInvoiceIssueDate.text = aInvoicesArray(UBound(aInvoicesArray), 10)
    txtInvoicePersonID.text = aInvoicesArray(1, 3)
    txtPersonDescription.text = aInvoicesArray(1, 4)
    txtInvoiceDeliveryPointID.text = aInvoicesArray(1, 6)
    txtDeliveryPointDescription.text = aInvoicesArray(1, 7)
    txtInvoicePaymentWayID.text = aInvoicesArray(1, 8)
    txtPaymentWayDescription.text = aInvoicesArray(1, 9)
    txtInvoiceRemarks.text = aInvoicesArray(1, 11)
    
End Function

Private Function AbortProcedure(blnStatus)
    
     If blnGridEditInProgress Then
        blnGridEditInProgress = False
        grdCommonTransactions.CancelEdit
        Exit Function
    End If
    
    If Not blnStatus Then
        If MyMsgBox(3, strAppTitle, strMessages(3), 2) Then
            blnStatus = False
            
            ClearFields txtInvoiceTrnID, txtInvoicePersonID, txtInvoiceCodeID, txtInvoicePrintExtraRemarksID, txtInvoiceDeliveryPointID, txtInvoicePaymentWayID, txtInvoiceInDate, txtInvoiceInTime, txtInvoiceIsInvoiced, txtInvoiceIsPrinted, txtCodeDetailsID, txtCodeHandID, txtCodeLastNo, txtVATStateID, txtCodeInventoryQty, txtCodeInventoryValue, txtCodeTransformID, mskCodeLastDate, txtCodePrinterID, mskCodeDetailLines, txtProfession, txtAddress, txtCity, txtTaxNo, txtPhones, txtTaxOfficeDescription, txtCodeDateCheckID, txtCodeIsCreditID, txtCodeDTable, txtCodeDTableCredit, grdCommonTransactions
            ClearFields mskInvoiceIssueDate, txtPersonDescription, txtCodeShortDescription, lblCodeDescription, txtInvoiceNo, txtInvoicePrintExtraRemarks, txtDeliveryPointDescription, txtPaymentWayDescription, txtInvoicePlates, txtInvoiceRemarks, txtInvoiceTransportReason, txtInvoiceTransportWay, txtInvoiceLoadingSite, txtInvoiceDestinationSite, txtCodeShortDescriptionCredit, txtInvoiceNoCredit
            ClearFields mskTotalQty, mskTotalPreDiscount, mskDiscount, mskTransDiscount, mskTotalRestAmount, mskExtraCharges, mskTotalVAT, mskTotalGross
            
            DisableFields mskInvoiceIssueDate, txtPersonDescription, txtCodeShortDescription, txtInvoiceNo, txtInvoicePrintExtraRemarks, txtDeliveryPointDescription, txtPaymentWayDescription, txtInvoicePlates, txtInvoiceRemarks, txtInvoiceTransportReason, txtInvoiceTransportWay, txtInvoiceLoadingSite, txtInvoiceDestinationSite, txtCodeShortDescriptionCredit, txtInvoiceNoCredit
            DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8), cmdIndex(9)
            DisableFields mskDiscount, mskTransDiscount, mskTotalRestAmount, mskExtraCharges, mskTotalVAT
            
            HideCredits
            
            UpdateButtons Me, 5, 1, 0, 0, IIf(CheckForLoadedForm("CommonTransactionsIndex"), 0, 1), 0, 1
        End If
        Exit Function
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Function CheckToEnableGrid()

    If mskInvoiceIssueDate.text <> "" And txtInvoicePersonID.text <> "" And txtInvoiceCodeID.text <> "" Then
        CheckToEnableGrid = True
    Else
        CheckToEnableGrid = False
    End If

End Function

Private Function DeleteInvoicesTrn()

    On Error GoTo ErrTrap
    
    Dim strSQL As String
    
    If blnError Then Exit Function
    
    If txtCodeHandID.text = "1" Or (txtCodeHandID.text = "0" And blnStatus) Then
        strSQL = "DELETE FROM InvoicesTrn WHERE InvoiceTrnID = " & Val(txtInvoiceTrnID.text)
        CommonDB.Execute (strSQL)
    End If
    
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
        
        ClearFields txtInvoiceTrnID, txtInvoicePersonID, txtInvoiceCodeID, txtInvoicePrintExtraRemarksID, txtInvoiceDeliveryPointID, txtInvoicePaymentWayID, txtInvoiceInDate, txtInvoiceInTime, txtInvoiceIsInvoiced, txtInvoiceIsPrinted, txtCodeDetailsID, txtCodeHandID, txtCodeLastNo, txtVATStateID, txtCodeInventoryQty, txtCodeInventoryValue, txtCodeTransformID, mskCodeLastDate, txtCodePrinterID, mskCodeDetailLines, txtProfession, txtAddress, txtCity, txtTaxNo, txtPhones, txtTaxOfficeDescription, grdCommonTransactions, txtCodeDateCheckID, txtCodeIsCreditID, txtCodeDTable, txtCodeDTableCredit
        ClearFields mskInvoiceIssueDate, txtPersonDescription, txtCodeShortDescription, lblCodeDescription, txtInvoiceNo, txtInvoicePrintExtraRemarks, txtDeliveryPointDescription, txtPaymentWayDescription, txtInvoicePlates, txtInvoiceRemarks, txtInvoiceTransportReason, txtInvoiceTransportWay, txtInvoiceLoadingSite, txtInvoiceDestinationSite, txtCodeShortDescriptionCredit, txtInvoiceNoCredit
        ClearFields mskTotalQty, mskTotalPreDiscount, mskDiscount, mskTransDiscount, mskTotalRestAmount, mskExtraCharges, mskTotalVAT, mskTotalGross
        
        DisableFields mskInvoiceIssueDate, txtPersonDescription, txtCodeShortDescription, txtInvoiceNo, txtInvoicePrintExtraRemarks, txtDeliveryPointDescription, txtPaymentWayDescription, txtInvoicePlates, txtInvoiceRemarks, txtInvoiceTransportReason, txtInvoiceTransportWay, txtInvoiceLoadingSite, txtInvoiceDestinationSite, txtCodeShortDescriptionCredit, txtInvoiceNoCredit
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8), cmdIndex(9)
        DisableFields mskDiscount, mskTransDiscount, mskTotalRestAmount, mskExtraCharges, mskTotalVAT
        
        HideCredits
        
        UpdateButtons Me, 5, 1, 0, 0, IIf(CheckForLoadedForm("CommonTransactionsIndex"), 0, 1), 0, 1
    Else
        Rollback
    End If
    
    Exit Function
    
ErrTrap:
    Rollback
    DeleteRecord = False
    DisplayErrorMessage True, Err.Description
    
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
    
    'Local μεταβλητές
    Dim lngIndex As Long
    Dim lngRow As Long
    
    'Qty
    Dim intThisQty As Integer
    Dim intLastQty As Integer
    
    'Recordsets
    Dim rstRecordset As Recordset
    Dim tmpRecordset As Recordset
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
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
    
    'Κύριο SQL
    strSQL = "SELECT InvoicesTrn.ItemID, Qty, UnitPrice, TotalNetPreDiscount, DiscPercent, DiscAmount, DiscAllow, TotalNetPostDiscount, VATPercent, VATAmount, TotalGross, ItemDescription, CategoryShortDescription, CategoryID, CategoryDescription, ManufacturerID, ManufacturerDescription, ItemBalance " _
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
            With grdCommonTransactions
                .AddRow , , , , , , , 1
                lngRow = .RowCount
                .CellValue(lngRow, "ItemID") = rstRecordset!ItemID
                .CellValue(lngRow, "ItemDescription") = rstRecordset!ItemDescription
                .CellValue(lngRow, "CategoryID") = rstRecordset!CategoryID
                .CellValue(lngRow, "CategoryShortDescription") = rstRecordset!CategoryShortDescription
                .CellValue(lngRow, "ManufacturerDescription") = rstRecordset!ManufacturerDescription
                .CellValue(lngRow, "Qty") = rstRecordset!Qty
                .CellValue(lngRow, "UnitPrice") = rstRecordset!UnitPrice
                .CellValue(lngRow, "TotalNetPreDiscount") = rstRecordset!TotalNetPreDiscount
                .CellValue(lngRow, "DiscPercent") = rstRecordset!DiscPercent
                .CellValue(lngRow, "DiscAmount") = rstRecordset!DiscAmount
                .CellValue(lngRow, "DiscAllow") = rstRecordset!DiscAllow
                .CellValue(lngRow, "TotalNetPostDiscount") = rstRecordset!TotalNetPostDiscount
                .CellValue(lngRow, "VATPercent") = rstRecordset!VATPercent
                .CellValue(lngRow, "VATAmount") = rstRecordset!VATAmount
                .CellValue(lngRow, "TotalGross") = rstRecordset!TotalGross
                
                lngItemID = .CellValue(lngRow, "ItemID")
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
                
                lngRow = lngRow + 1
            
            End With
            .MoveNext
        Wend
    End With
    
    'Τελικές ενέργειες
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
    
    Dim lngRow As Long
    
    blnStatus = True
    
    ClearFields txtInvoiceTrnID, txtInvoicePersonID, txtInvoiceCodeID, txtInvoicePrintExtraRemarksID, txtInvoiceDeliveryPointID, txtInvoicePaymentWayID, txtInvoiceInDate, txtInvoiceInTime, txtInvoiceIsInvoiced, txtInvoiceIsPrinted, txtCodeDetailsID, txtCodeHandID, txtCodeLastNo, txtVATStateID, txtCodeInventoryQty, txtCodeInventoryValue, txtCodeTransformID, mskCodeLastDate, txtCodePrinterID, mskCodeDetailLines, txtProfession, txtAddress, txtCity, txtTaxNo, txtPhones, txtTaxOfficeDescription, grdCommonTransactions, txtCodeDateCheckID, txtCodeIsCreditID, txtCodeDTable, txtCodeDTableCredit
    ClearFields mskInvoiceIssueDate, txtPersonDescription, txtCodeShortDescription, lblCodeDescription, txtInvoiceNo, txtInvoicePrintExtraRemarks, txtDeliveryPointDescription, txtPaymentWayDescription, txtInvoicePlates, txtInvoiceRemarks, txtInvoiceTransportReason, txtInvoiceTransportWay, txtInvoiceLoadingSite, txtInvoiceDestinationSite, txtCodeShortDescriptionCredit, txtInvoiceNoCredit
    ClearFields mskTotalQty, mskTotalPreDiscount, mskDiscount, mskTransDiscount, mskTotalRestAmount, mskExtraCharges, mskTotalVAT, mskTotalGross
    
    EnableFields mskInvoiceIssueDate, txtPersonDescription, txtCodeShortDescription, txtInvoiceNo, txtInvoicePrintExtraRemarks, txtDeliveryPointDescription, txtPaymentWayDescription, txtInvoicePlates, txtInvoiceRemarks, txtInvoiceTransportReason, txtInvoiceTransportWay, txtInvoiceLoadingSite, txtInvoiceDestinationSite, txtCodeShortDescriptionCredit, txtInvoiceNoCredit
    EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8)
    EnableFields mskDiscount, mskTransDiscount, mskTotalRestAmount, mskExtraCharges, mskTotalVAT
    
    ReDim aInvoicesArray(1, 3)
    
    CustomizeGrid grdCommonTransactions
    EnableTabStop grdCommonTransactions
    
    EnableOrDisableSpecialFields
    
    InitializeFields IIf(txtRefersTo.text = "2", mskInvoiceIssueDate, ""), mskTotalQty, mskTotalPreDiscount, mskDiscount, mskTransDiscount, mskTotalRestAmount, mskExtraCharges, mskTotalVAT, mskTotalGross
    ColorizeRowsWhenItemIsNotGiven 0
    FillCellWithSomething grdCommonTransactions, 0, 0, "5,6,7,8,9,10,12,13,14,15,16"
    
    UpdateButtons Me, 5, 0, 1, 0, 0, 1, 0
    
    mskInvoiceIssueDate.SetFocus
    
End Function

Private Function PrintInvoice()

    'Αν είναι χειρόγραφο, βγαίνω
    If txtCodeHandID.text = "1" Then Exit Function
    
    If PrintRecords(Me, "Print", False, "PrinterPrintsInvoicesID", txtCodePrinterID.text, Val(txtInvoiceTrnID.text)) Then
        blnError = True
    Else
        blnError = False
    End If

End Function

Function CreateUnicodeFile(myPrinterType, myEAFDSSString, myInvoiceHeight, myDetailLines, myTopMargin, myLeftMargin)

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
    
    On Error GoTo ErrTrap
    
    Dim lngRow As Long
    Dim intDetailLines As Integer
    
    Open strUnicodeFile For Output As #1
    InitReport myPrinterType, myEAFDSSString, myInvoiceHeight
    GoSub PrintInvoiceHeadings
    
    With grdCommonTransactions
        For lngRow = 1 To grdCommonTransactions.RowCount
            If .CellValue(lngRow, "CategoryID") <> "" And .CellValue(lngRow, "ItemID") <> "" Then
                intDetailLines = intDetailLines + 1
                Print #1, ItemDescriptionAndManufacturer(.CellText(lngRow, "ItemDescription"), .CellText(lngRow, "ManufacturerDescription")); Tab(48); "TEM"; Tab(60 - Len(.CellText(lngRow, "Qty"))); .CellText(lngRow, "Qty"); _
                Tab(74 - Len(format(.CellText(lngRow, "UnitPrice"), "#,##0.00"))); format(.CellText(lngRow, "UnitPrice"), "#,##0.00"); _
                Tab(89 - Len(format(.CellText(lngRow, "TotalNetPreDiscount"), "#,##0.00"))); format(.CellText(lngRow, "TotalNetPreDiscount"), "#,##0.00"); _
                Tab(96 - Len(format(.CellText(lngRow, "DiscPercent"), "#,##0.00"))); format(.CellText(lngRow, "DiscPercent"), "#,##0.00"); _
                Tab(107 - Len(format(.CellText(lngRow, "DiscAmount"), "#,##0.00"))); format(.CellText(lngRow, "DiscAmount"), "#,##0.00"); _
                Tab(119 - Len(format(.CellText(lngRow, "TotalNetPostDiscount"), "#,##0.00"))); format(.CellText(lngRow, "TotalNetPostDiscount"), "#,##0.00"); _
                Tab(123 - Len(format(.CellValue(lngRow, "VATPercent"), "#0"))); format(.CellValue(lngRow, "VATPercent"), "#0,00"); _
                Tab(136 - Len(format(.CellText(lngRow, "VATAmount"), "#,##0.00"))); format(.CellValue(lngRow, "VATAmount"), "#,##0.00")
            End If
        Next lngRow
    End With
    
    'For intDetailLines = intDetailLines To myDetailLines - 13
    '    Print #1, ""
    'Next intDetailLines
    
    'Print #1, "ΓΡΑΜΜΗ 1"
    'Print #1, "ΓΡΑΜΜΗ 2"
    
    For intDetailLines = intDetailLines To myDetailLines - 12
        Print #1, ""
    Next intDetailLines
    
    '27+69 Bold 27+70 Cancel
    '27+71 Double 27+72 Cancel
    
    If txtInvoicePrintExtraRemarksID.text = "1" Then
        Print #1, Chr(14) + strInvoiceExtraRemarksA + Chr(20)
    Else
        Print #1, ""
    End If
    Print #1, strInvoiceExtraRemarksB
    Print #1, ""
    Print #1, ""
    
    'If blnPrintBalance Then Print #1, Tab(35 - Len(format(curPreviousBalance, "#,##0.00"))); format(curPreviousBalance, "#,##0.00");
    'If blnPrintBalance Then Print #1, Tab(35 - Len(format(curNewBalance, "#,##0.00"))); format(curNewBalance, "#,##0.00");
    
    Print #1, Tab(136 - Len(format(mskTotalPreDiscount.text, "#,##0.00"))); format(mskTotalPreDiscount.text, "#,##0.00")
    Print #1, Tab(136 - Len(format(mskDiscount.text, "#,##0.00"))); format(mskDiscount.text, "#,##0.00")
    Print #1, Tab(45); format(mskTotalRestAmount.text, "#,##0.00"); Tab(65); CStr(curExtraChargesVATPercent); Tab(75); format(mskTotalVAT.text, "#,##0.00"); Tab(136 - Len(format(mskTotalRestAmount.text, "#,##0.00"))); format(mskTotalRestAmount.text, "#,##0.00")
    Print #1, Tab(136 - Len(format(mskTotalVAT.text, "#,##0.00"))); format(mskTotalVAT.text, "#,##0.00")
    Print #1, ""
    Print #1, Tab(136 - Len(format(mskTotalGross.text, "#,##0.00"))); Chr(27) + Chr(71) + format(mskTotalGross.text, "#,##0.00") + Chr(27) + Chr(72)
    
    Print #1, Space(13) & Left(txtInvoiceRemarks.text, 60)
    Print #1, FullNumber(mskTotalGross.text)
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    
    Close #1
    
    Printer.EndDoc
    
    CreateUnicodeFile = strUnicodeFile
    
    Exit Function

PrintInvoiceHeadings:
    
    For myTopMargin = 1 To myTopMargin - 1
        Print #1, ""
    Next myTopMargin

    Print #1, Tab(11); lblCodeDescription.Caption; Tab(95 - (Len(txtInvoiceNo.text) / 2)); txtInvoiceNo.text; Tab(107); mskInvoiceIssueDate.text; Tab(128);: 'If blnPrintHour = True Then Print #1, strTime
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, Tab(11); txtInvoicePersonID.text; Tab(40); Left(txtTaxOfficeDescription, 20); Tab(86); Left(txtInvoiceTransportReason.text, 17); Tab(121); IIf(blnUseNewInvoiceForm, "", Left(txtInvoicePlates.text, 16))
    Print #1, Tab(11); txtPersonDescription.text; Tab(86); Left(txtInvoiceLoadingSite.text, 17)
    Print #1, Tab(11); txtProfession.text; Tab(86); Left(txtInvoiceDestinationSite.text, 17)
    Print #1, Tab(11); txtAddress.text; Tab(86); Left(txtInvoiceTransportWay.text, 17)
    Print #1, Tab(11); txtCity.text; Tab(86); Left(txtPaymentWayDescription.text, 40)
    Print #1, Tab(11); txtTaxNo.text; Tab(40); Left(txtPhones.text, 20), Tab(86); IIf(blnUseNewInvoiceForm, Left(txtInvoicePlates.text, 16), "")
    
    Print #1, ""
    Print #1, ""
    Print #1, ""
    
    intDetailLines = 15
    
    Return
    
ErrTrap:
    Close #1
    CreateUnicodeFile = "Error"
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function SaveInvoicesTrn()

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
    
    Dim lngRow As Long
    
    If blnError Then Exit Function
    
    If txtCodeHandID.text = "1" Or (txtCodeHandID.text = "0" And blnStatus) Then
    
        With grdCommonTransactions
            For lngRow = 1 To .RowCount
                If .CellValue(lngRow, "ItemID") <> "" Then
                    If Not MainSaveRecord("CommonDB", "InvoicesTrn", True, strAppTitle, "InvoiceTrnID", txtInvoiceTrnID.text, _
                        .CellValue(lngRow, "ItemID"), _
                        .CellValue(lngRow, "Qty"), _
                        .CellValue(lngRow, "UnitPrice"), _
                        .CellValue(lngRow, "TotalNetPreDiscount"), _
                        .CellValue(lngRow, "DiscPercent"), _
                        .CellValue(lngRow, "DiscAmount"), _
                        .CellValue(lngRow, "DiscAllow"), _
                        .CellValue(lngRow, "TotalNetPostDiscount"), _
                        .CellValue(lngRow, "VATPercent"), _
                        .CellValue(lngRow, "VATAmount"), _
                        .CellValue(lngRow, "TotalGross"), _
                        lngTrnID) <> 0 Then
                        blnError = True
                    End If
                End If
            Next lngRow
        End With
    
        SaveInvoicesTrn = True
    
    End If
    
End Function

Private Function SaveInvoice()

    Dim lngRow As Long
    
    If blnError Then Exit Function
    
    ValidateFields
    
    'Αν το παραστατικό είναι χειρόγραφο 'Η
    'Αν είναι μηχανογραφικό και γίνεται έλεγχος ημερομηνίας και είμαι σε νέα εγγραφή 'Η
    'Αν είναι μηχανογραφικό και δεν γίνεται έλεγχος ημερομηνίας
    
    If txtCodeHandID.text = "1" Or (txtCodeHandID.text = "0" And txtCodeDateCheckID = "1" And blnStatus) Or (txtCodeHandID.text = "0" And txtCodeDateCheckID = "0") Then
        
        lngTrnID = IIf(txtInvoiceTrnID.text = "", AddOneToTheLastRecord, txtInvoiceTrnID.text)
        txtInvoiceTrnID.text = lngTrnID
        txtInvoiceIsInvoiced.text = IIf(blnStatus, txtCodeTransformID.text, txtInvoiceIsInvoiced.text) 'Τιμολογημένο 0 = Δεν χρειάζεται, 1= Εκκρεμεί τιμολόγηση, 2 = Τιμολογημένο
        txtInvoiceIsPrinted.text = IIf(txtCodeHandID.text = "1", "0", "1") 'Εκτυπωμένο 0 = Οχι, 1 = Ναι (Αναλόγως το παραστατικό)
        
        If Not MainSaveRecord("CommonDB", "Invoices", blnStatus, strAppTitle, "TrnID", _
            txtInvoiceTrnID.text, _
            mskInvoiceIssueDate.text, txtInvoiceNo.text, txtInvoiceCodeID.text, Val(txtRefersTo.text), _
            mskTotalQty.text, mskTotalPreDiscount.text, mskDiscount.text, mskTransDiscount.text, mskTotalRestAmount.text, mskExtraCharges.text, mskTotalVAT.text, mskTotalGross.text, _
            lngTrnID, _
            txtInvoiceRemarks.text, _
            txtInvoicePlates.text, _
            txtInvoicePaymentWayID.text, _
            txtInvoicePersonID.text, _
            txtInvoiceIsInvoiced.text, _
            txtInvoiceIsPrinted.text, _
            IIf(blnStatus, Date, txtInvoiceInDate.text), _
            IIf(blnStatus, Time, txtInvoiceInTime.text), _
            txtInvoiceTransportReason.text, _
            txtInvoiceTransportWay.text, _
            txtInvoiceLoadingSite.text, _
            txtInvoiceDestinationSite.text, _
            txtInvoiceDeliveryPointID.text, _
            IIf(txtInvoicePrintExtraRemarksID.text <> "", txtInvoicePrintExtraRemarksID.text, "0"), _
            strCurrentUser) <> 0 Then
            blnError = True
        End If
    
    End If
    
End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    blnError = False
    
    If txtCodeHandID.text = "1" Or (txtCodeHandID.text = "0" And blnStatus) Then BeginTrans
    
    'DeleteInvoicesTrn
    'SaveInvoice
    'SaveInvoicesTrn
    SaveELine
    'UpdateCodes
    'UpdateItemsWithNewBalance
    'TransformInvoices
    'PrintInvoice
        
    'If Not blnError Then
    '    If txtCodeHandID.text = "1" Or (txtCodeHandID.text = "0" And blnStatus) Then CommitTrans
    '
    '    ClearFields txtInvoiceTrnID, txtInvoicePersonID, txtInvoiceCodeID, txtInvoicePrintExtraRemarksID, txtInvoiceDeliveryPointID, txtInvoicePaymentWayID, txtInvoiceInDate, txtInvoiceInTime, txtInvoiceIsInvoiced, txtInvoiceIsPrinted, txtCodeDetailsID, txtCodeHandID, txtCodeLastNo, txtVATStateID, txtCodeInventoryQty, txtCodeInventoryValue, txtCodeTransformID, mskCodeLastDate, txtCodePrinterID, mskCodeDetailLines, txtProfession, txtAddress, txtCity, txtTaxNo, txtPhones, txtTaxOfficeDescription, grdCommonTransactions, txtCodeDateCheckID, txtCodeIsCreditID, txtCodeDTable, txtCodeDTableCredit
    '    ClearFields mskInvoiceIssueDate, txtPersonDescription, txtCodeShortDescription, lblCodeDescription, txtInvoiceNo, txtInvoicePrintExtraRemarks, txtDeliveryPointDescription, txtPaymentWayDescription, txtInvoicePlates, txtInvoiceRemarks, txtInvoiceTransportReason, txtInvoiceTransportWay, txtInvoiceLoadingSite, txtInvoiceDestinationSite, txtCodeShortDescriptionCredit, txtInvoiceNoCredit
    '    ClearFields mskTotalQty, mskTotalPreDiscount, mskDiscount, mskTransDiscount, mskTotalRestAmount, mskExtraCharges, mskTotalVAT, mskTotalGross
    '
    '    DisableFields mskInvoiceIssueDate, txtPersonDescription, txtCodeShortDescription, txtInvoiceNo, txtInvoicePrintExtraRemarks, txtDeliveryPointDescription, txtPaymentWayDescription, txtInvoicePlates, txtInvoiceRemarks, txtInvoiceTransportReason, txtInvoiceTransportWay, txtInvoiceLoadingSite, txtInvoiceDestinationSite, txtCodeShortDescriptionCredit, txtInvoiceNoCredit
    '    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8), cmdIndex(9)
    '    DisableFields mskDiscount, mskTransDiscount, mskTotalRestAmount, mskExtraCharges, mskTotalVAT
    '
    '    HideCredits
    '
    '    UpdateButtons Me, 5, 1, 0, 0, IIf(CheckForLoadedForm("CommonTransactionsIndex"), 0, 1), 0, 1
    'Else
    '    If txtCodeHandID.text = "1" Or (txtCodeHandID.text = "0" And blnStatus) Then Rollback
    'End If
    
End Function

Function FindInvoicesWithTrnID(myInvoiceTrnID, myWindowTitle, myTable, myRefersTo)

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
    txtTable.text = myTable
    txtRefersTo.text = myRefersTo
    blnStatus = False
    ReDim aInvoicesArray(1, 3)
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Κύριο SQL
    strSQL = "SELECT InvoiceIssueDate, InvoiceNo, InvoiceCodeID, InvoiceRefersToID, InvoiceQty, InvoiceNet, InvoicePercentDiscount, InvoiceAmountDiscount, InvoiceRestAmount, InvoiceVATAmount, InvoiceGrossAmount, InvoiceTrnID, InvoiceRemarks, InvoicePlates, InvoicePaymentWayID, InvoicePersonID, InvoiceIsInvoiced, InvoiceIsPrinted, InvoiceInDate, InvoiceInTime, InvoiceExtraChargesAmount, InvoiceTransportReason, InvoiceTransportWay, InvoiceLoadingSite, InvoiceDestinationSite, InvoiceDeliveryPointID, InvoicePrintExtraRemarksID " _
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
    
    'Ενημερώνω τα πεδία της φόρμας
    With rstRecordset
        'Ημερομηνία
        mskInvoiceIssueDate.text = format(!InvoiceIssueDate, "dd/mm/yyyy")
        'Συναλλασόμενος
        FindPersonDetails !InvoicePersonID
        'Παραστατικό
        txtInvoiceCodeID.text = !InvoiceCodeID
        Set tmpRecordset = CheckForMatch("CommonDB", txtInvoiceCodeID.text, "Codes", "CodeID", "Numeric", 0, 1)
        txtInvoiceCodeID.text = tmpRecordset.Fields(0)
        txtCodeShortDescription.text = tmpRecordset.Fields(1)
        lblCodeDescription.Caption = tmpRecordset.Fields(2)
        txtCodeDetailsID.text = tmpRecordset.Fields(9)
        txtCodeHandID.text = tmpRecordset.Fields(10)
        txtCodeInventoryQty.text = tmpRecordset.Fields(4)
        txtCodeInventoryValue.text = tmpRecordset.Fields(5)
        txtCodeTransformID.text = tmpRecordset.Fields(12)
        txtCodeLastNo.text = tmpRecordset.Fields(14)
        mskCodeLastDate.text = tmpRecordset.Fields(15)
        txtCodePrinterID.text = tmpRecordset.Fields(11)
        mskCodeDetailLines.text = tmpRecordset.Fields(16)
        txtCodeDateCheckID.text = tmpRecordset.Fields(17)
        txtInvoiceDeliveryPointID.text = !InvoiceDeliveryPointID
        txtInvoicePrintExtraRemarksID.text = !InvoicePrintExtraRemarksID
        'Νο παραστατικού
        txtInvoiceNo.text = !InvoiceNo
        'Τυπώνει επιπλέον παρατηρήσεις
        Set tmpRecordset = CheckForMatch("CommonDB", txtInvoicePrintExtraRemarksID.text, "YesOrNo", "YesNoID", "String", 0, 1)
        txtInvoicePrintExtraRemarksID.text = tmpRecordset.Fields(0)
        txtInvoicePrintExtraRemarks.text = tmpRecordset.Fields(1)
        'Τόπος παραλαβής
        Set tmpRecordset = CheckForMatch("CommonDB", txtInvoiceDeliveryPointID.text, "DeliveryPoints", "DeliveryPointID", "Numeric", 0, 1)
        txtInvoiceDeliveryPointID.text = tmpRecordset.Fields(0)
        txtDeliveryPointDescription.text = tmpRecordset.Fields(1)
        'Τρόπος πληρωμής
        txtInvoicePaymentWayID.text = !InvoicePaymentWayID
        Set tmpRecordset = CheckForMatch("CommonDB", txtInvoicePaymentWayID.text, "PaymentWays", "PaymentWayID", "Numeric", 0, 1)
        txtInvoicePaymentWayID.text = tmpRecordset.Fields(0)
        txtPaymentWayDescription.text = tmpRecordset.Fields(1)
        'Αρ. κυκλοφορίας
        txtInvoicePlates.text = !InvoicePlates
        'Παρατηρήσεις
        txtInvoiceRemarks.text = !InvoiceRemarks
        'Λοιπά στοιχεία
        txtInvoiceTransportReason.text = !InvoiceTransportReason
        txtInvoiceTransportWay.text = !InvoiceTransportWay
        txtInvoiceLoadingSite.text = !InvoiceLoadingSite
        txtInvoiceDestinationSite.text = !InvoiceDestinationSite
        'Σύνολα
        mskTotalQty.text = format(!InvoiceQty, "#,##0")
        mskTotalPreDiscount.text = format(!InvoiceNet, "#,##0.00")
        mskDiscount.text = format(!InvoicePercentDiscount, "#,##0.00")
        mskTransDiscount.text = format(!InvoiceAmountDiscount, "#,##0.00")
        mskTotalRestAmount.text = format(!InvoiceRestAmount, "#,##0.00")
        mskExtraCharges.text = format(!InvoiceExtraChargesAmount, "#,##0.00")
        mskTotalVAT.text = format(!InvoiceVATAmount, "#,##0.00")
        mskTotalGross.text = format(!InvoiceGrossAmount, "#,##0.00")
        'Βοηθητικά
        txtInvoiceTrnID.text = !InvoiceTrnID
        txtInvoiceIsInvoiced.text = !InvoiceIsInvoiced
        txtInvoiceIsPrinted.text = !InvoiceIsPrinted
        txtInvoiceInDate.text = format(!InvoiceInDate, "dd/mm/yy")
        txtInvoiceInTime.text = format(!InvoiceInTime, "hh:mm")
        txtRefersTo.text = !InvoiceRefersToID
    End With
    
    CustomizeGrid grdCommonTransactions
    EnableTabStop grdCommonTransactions
    
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
        .txtTable.text = txtTable.text
        .Tag = "True"
        DisableFields .txtCategoryShortDescription, .txtManufacturerDescription, .txtItemDescription, .cmdIndex(0), .cmdIndex(1), .cmdIndex(2)
        .Show 1, Me
    End With

End Function

Function ShowOrHideDetailsFrame()

    If txtRefersTo.text = "1" Then
        frmDetails.Visible = False
        grdCommonTransactions.Height = Me.Height - 6190
   Else
        frmDetails.Visible = True
        grdCommonTransactions.Height = Me.Height - 6280 - frmDetails.Height
    End If
   
End Function



Private Function ShowReport()

    With CommonTransactionsIndex
        .lblTitle.Caption = WindowTitle(lblTitle.Caption)
        .txtTable.text = txtTable.text
        .txtRefersTo.text = txtRefersTo.text
        .Tag = "True"
        .Show 1, Me
    End With

End Function

Private Function UpdateCodes()

    Dim rsCodes As Recordset
    
    Set rsCodes = CommonDB.OpenRecordset("Codes")
    
    If txtCodeHandID.text = "1" Or (txtCodeHandID.text = "0" And blnStatus) Then
    
        If txtRefersTo.text = "2" And txtCodeHandID.text = "0" Then
            With rsCodes
                .Index = "ID"
                .Seek "=", Val(txtInvoiceCodeID.text)
                If !CodeHandID = 0 Then
                    .Edit
                    !CodeLastNo = txtInvoiceNo.text
                    !CodeLastDate = mskInvoiceIssueDate.text
                    .Update
                End If
            End With
        End If
    
    End If
    
    rsCodes.Close

End Function

Private Function UpdateColTags()

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

    'Αγορές - όχι τιμολόγηση Δ.Α.
    If txtRefersTo.text = "1" And aInvoicesArray(1, 1) = "" Then
        grdCommonTransactions.ColTag("CategoryShortDescription") = "Y"
        grdCommonTransactions.ColTag("ItemDescription") = "Y"
        grdCommonTransactions.ColTag("Qty") = IIf(txtCodeInventoryQty.text = "", "N", "Y")
        grdCommonTransactions.ColTag("UnitPrice") = IIf(txtCodeInventoryValue.text = "", "N", "Y")
        grdCommonTransactions.ColTag("DiscPercent") = IIf(txtCodeInventoryValue.text = "", "N", "Y")
        grdCommonTransactions.ColTag("DiscAmount") = IIf(txtCodeInventoryValue.text = "", "N", "Y")
        grdCommonTransactions.ColTag("DiscAllow") = "N"
        grdCommonTransactions.ColTag("TotalGross") = "N"
    End If
    
    'Αγορές - τιμολόγηση Δ.Α.
    If txtRefersTo.text = "1" And aInvoicesArray(1, 2) <> "" Then
        grdCommonTransactions.ColTag("CategoryShortDescription") = "N"
        grdCommonTransactions.ColTag("ItemDescription") = "N"
        grdCommonTransactions.ColTag("Qty") = "N"
        grdCommonTransactions.ColTag("UnitPrice") = "Y"
        grdCommonTransactions.ColTag("DiscPercent") = "Y"
        grdCommonTransactions.ColTag("DiscAmount") = "Y"
        grdCommonTransactions.ColTag("DiscAllow") = "N"
        grdCommonTransactions.ColTag("TotalGross") = "N"
    End If
    
    'Πωλήσεις
    If txtRefersTo.text = "2" Then
        grdCommonTransactions.ColTag("Qty") = IIf(txtCodeInventoryQty.text = "", "N", "Y")
        grdCommonTransactions.ColTag("UnitPrice") = "N"
        grdCommonTransactions.ColTag("DiscPercent") = "N"
        grdCommonTransactions.ColTag("DiscAmount") = "N"
        grdCommonTransactions.ColTag("DiscAllow") = "N"
        grdCommonTransactions.ColTag("TotalGross") = IIf(txtCodeInventoryValue.text = "", "N", "Y")
    End If

End Function

Private Function UpdateFieldsWithDetails()

    txtInvoiceTransportReason.text = strTransportReason
    txtInvoiceLoadingSite.text = strLoadingSite
    txtInvoiceTransportWay.text = strTransportWay
    txtInvoiceDestinationSite.text = strDestinationSite

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
    
    Dim lngRow As Long
    
    Dim intQty As Integer
    Dim lngItemID As Long
    Dim intLastQty As Integer
    Dim intThisQty As Integer
    Dim intNewQty As Integer
    
    Dim rsItems As Recordset
    
    If blnError Then Exit Function
    
    Set rsItems = CommonDB.OpenRecordset("Items")
    
    With grdCommonTransactions
        For lngRow = 1 To .RowCount
            If .CellValue(lngRow, "ItemID") <> "" Then
                
                lngItemID = .CellValue(lngRow, "ItemID")
                intLastQty = IIf(.CellValue(lngRow, "LastQty") <> "", .CellValue(lngRow, "LastQty"), 0)
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
            AbortProcedure False
        Case 5
            AbortProcedure True
    End Select

End Sub

Private Sub cmdIndex_Click(Index As Integer)

    On Error GoTo ErrTrap
    
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    Dim strFieldQuery As String
    
    Select Case Index
        Case 0
            'Συναλλασόμενος
            If txtPersonDescription.text = "" Then Exit Sub
            'Αν έχω δώσει δύο αστεράκια και είμαι στις πωλήσεις
            If Left(txtPersonDescription.text, 2) = "**" And txtRefersTo.text = "2" Then
                'Αναζήτηση με βάση τις πινακίδες
                strFieldQuery = "Invoices.InvoiceRefersToID = 2 AND InStr(Invoices.InvoicePlates,'" & Mid(txtPersonDescription.text, 3, Len(txtPersonDescription.text)) & "')"
                Set tmpRecordset = NewCheckForMatch("CommonDB", "Customers.ID, Customers.Description, Customers.TaxNo, Invoices.InvoicePlates, Customers.Active", _
                    "Invoices", _
                    "INNER JOIN Customers ON Invoices.InvoicePersonID = Customers.ID", strFieldQuery, "GROUP BY Customers.ID, Customers.Description,Customers.TaxNo, Customers.Active, Invoices.InvoicePlates", "Invoices.InvoicePlates")
                If tmpRecordset.RecordCount > 0 Then
                    tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 5, 0, 1, 2, 3, 4, "ID", "Επωνυμία", "Α.Φ.Μ.", "Αρ. κυκλοφορίας", "Ε", 0, 50, 15, 15, 0, 1, 0, 1, 1, 1, "Persons")
                    txtInvoicePlates.text = tmpTableData.strThreeField
                End If
            End If
            'Αναζήτηση με επωνυμία ή περιγραφή όπως πάντα!
            If Left(txtPersonDescription.text, 2) <> "**" Then
                Set tmpRecordset = CheckForMatch("CommonDB", txtPersonDescription.text, txtTable.text, IIf(IsNumeric(txtPersonDescription.text), "TaxNo", "Description"), "String", 1, 2)
                If tmpRecordset.RecordCount > 0 Then
                    tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 4, 0, 1, 2, 13, "ID", "Περιγραφή", "Α.Φ.Μ.", "Ε", 0, 50, 15, 0, 1, 0, 1, 1, "Persons")
                End If
            End If
            If tmpTableData.strCode <> "" Then
                FindPersonDetails tmpTableData.strCode
                Dim eLine As String
                eLine = CreateELine
            End If
        Case 1
            'Συναλλασόμενος
            With Persons
                .txtTable.text = txtTable.text
                .txtRefersTo.text = txtRefersTo.text
                .lblTitle.Caption = IIf(txtRefersTo.text = "1", "Προμηθευτές", "Πελάτες")
                .Tag = "True"
                .Show 1, Me
            End With
        Case 2
            'Παραστατικό
            If txtCodeShortDescription.text = "" Then Exit Sub
            Set tmpRecordset = CheckForMatch("CommonDB", txtCodeShortDescription.text, "Codes", "CodeShortDescription", "String", txtRefersTo.text, "CodeDescription")
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 18, _
                0, 1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 16, 17, 18, 19, _
                "ID", "Συντ", "Περιγραφή", "Ποσότητες", "Αξίες", "Πελάτες", "Προμηθευτές", "Σειρά", "Λοιπά στοιχεία", "Χειρόγραφο", "Εκτυπωτής", "Πίν. Δ.", "Ακυρωτικό", "Μετασχηματίζεται", "Τελ. Νο", "Ημερομηνία", "Γραμμές", "Ελεγχος ημερομηνίας", _
                5, 6, 40, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 10, 5, 10, _
                1, 1, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
            txtInvoiceCodeID.text = tmpTableData.strCode
            txtCodeShortDescription.text = tmpTableData.strOneField
            lblCodeDescription.Caption = tmpTableData.strTwoField & IIf(tmpTableData.strSevenField <> "", " ( ΣΕΙΡΑ " & tmpTableData.strSevenField & " )", "")
            txtCodeInventoryQty.text = tmpTableData.strThreeField
            txtCodeInventoryValue.text = tmpTableData.strFourField
            txtCodeDetailsID.text = tmpTableData.strEightField
            txtCodeHandID.text = tmpTableData.strNineField
            txtCodePrinterID.text = tmpTableData.strTenField
            txtCodeDTable.text = tmpTableData.strElevenField
            txtCodeIsCreditID.text = tmpTableData.strTwelveField
            txtCodeTransformID.text = tmpTableData.strThirteenField
            txtCodeLastNo.text = tmpTableData.strFourteenField
            mskCodeLastDate.text = tmpTableData.strFifteenField
            mskCodeDetailLines.text = tmpTableData.strSixteenField
            txtCodeDateCheckID.text = tmpTableData.strSeventeenField
            
            If tmpRecordset.RecordCount <> 0 And txtInvoiceCodeID.text <> "" Then
                AddGridLines
                If txtRefersTo.text = "2" Then txtInvoiceNo.text = Val(txtCodeLastNo.text) + 1 'Αν είναι πώληση, αυξάνω τον αριθμό παραστατικού κατά 1
                If txtCodeHandID.text = "1" Then txtInvoiceNo.Locked = False Else txtInvoiceNo.Locked = True 'Αν είναι χειρόγραφο, επιτρέπω την αλλαγή του αριθμού
                If txtCodeDetailsID.text = "1" Then
                    UpdateFieldsWithDetails
                    EnableFields txtInvoiceTransportReason, txtInvoiceTransportWay, txtInvoiceLoadingSite, txtInvoiceDestinationSite
                End If
                CreateELine
            End If
            If txtCodeIsCreditID.text = "1" Then
                ShowCredits
            Else
                HideCredits
            End If
        Case 3
            'Παραστατικό
            With UtilsCodes
                .Tag = "True"
                .txtRefersTo.text = txtRefersTo.text
                .Show 1, Me
            End With
        Case 4
            'Εξτρα παρατηρήσεις
            Set tmpRecordset = CheckForMatch("CommonDB", txtInvoicePrintExtraRemarks.text, "YesOrNo", "YesNoDescription", "String", 1, 2)
            tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtInvoicePrintExtraRemarksID.text = tmpTableData.strCode
            txtInvoicePrintExtraRemarks.text = tmpTableData.strOneField
        Case 5
            'Τόπος παραλαβής
            If txtDeliveryPointDescription.text = "" Then Exit Sub
            With UtilsDeliveryPoints
                Set tmpRecordset = CheckForMatch("CommonDB", txtDeliveryPointDescription.text, "DeliveryPoints", "DeliveryPointDescription", "String", 1, 2)
                tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtInvoiceDeliveryPointID.text = tmpTableData.strCode
                txtDeliveryPointDescription.text = tmpTableData.strOneField
            End With
        Case 6
            'Τόπος παραλαβής
            With UtilsDeliveryPoints
                .Tag = "True"
                .Show 1, Me
            End With
        Case 7
            'Τρόπος πληρωμής
            If txtPaymentWayDescription.text = "" Then Exit Sub
            With UtilsPaymentWays
                Set tmpRecordset = CheckForMatch("CommonDB", txtPaymentWayDescription.text, "PaymentWays", "PaymentWayDescription", "String", 1, 2)
                tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtInvoicePaymentWayID.text = tmpTableData.strCode
                txtPaymentWayDescription.text = tmpTableData.strOneField
            End With
        Case 8
            'Τρόπος πληρωμής
            With UtilsPaymentWays
                .Tag = "True"
                .Show 1, Me
            End With
        Case 9
            'Καρτέλα συναλλασόμενου
            ShowPersonLedger _
                txtInvoicePersonID.text, _
                txtPersonDescription.text, _
                IIf(txtRefersTo.text = "1", "Καρτέλα προμηθευτή", "Καρτέλα πελάτη"), _
                txtTable.text, _
                IIf(txtRefersTo.text = "1", "Customers", "Suppliers"), _
                IIf(txtRefersTo.text = "1", "3", "4")
            Case 10
                'Παραστατικό που ακυρώνεται
                If txtCodeShortDescriptionCredit.text = "" Then Exit Sub
                Set tmpRecordset = CheckForMatch("CommonDB", txtCodeShortDescriptionCredit.text, "Codes", "CodeShortDescription", "String", txtRefersTo.text, "CodeDescription")
                tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 4, _
                    0, 1, 2, 12, _
                    "ID", "Συντ.", "Περιγραφή", "Κωδ. πίνακα Δ", _
                    0, 6, 40, 40, _
                    1, 1, 0, 1)
                txtInvoiceCodeCreditID.text = tmpTableData.strCode
                txtCodeShortDescriptionCredit.text = tmpTableData.strOneField
                txtCodeDTableCredit.text = tmpTableData.strThreeField
                CreateELine
    End Select
    
    Exit Sub
    
ErrTrap:
    Exit Sub
    '                "ID", "Συντ", "Περιγραφή", "Ποσότητα", "Αξία", "Πελάτης", "Προμηθευτής", "Σειρά", "Λοιπά στοιχεία", "Χειρόγραφο", "Εκτυπωτής", "Πίν. Δ", "Ακυρωτικό", "Μετασχηματίζεται", "Νο", "Ημερομηνία", "Γραμμές", "Ελεγχος ημερομηνίας", _

End Sub

Private Function SaveELine()

    Open strReportsPathName + "eLine.txt" For Append As #1
    
    Print #1, eLine
    
    Close #1
    
End Function


Private Sub Form_Activate()

    If Me.Tag = "True" Then
        Me.Tag = "False"
        ShowOrHideDetailsFrame
    End If
    
    'AddDummyLines grdCommonTransactions, 2, 5, 2, 50, 40, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10
    'grdCommonTransactions.Enabled = True

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
        Case vbKeyEscape
            If cmdButton(4).Enabled Then cmdButton_Click 4: Exit Function
            If cmdButton(5).Enabled Then cmdButton_Click 5: Exit Function
        Case vbKeyF12 And CtrlDown = 4
            ToggleInfoPanel Me
        Case vbKeyE And CtrlDown = 4 And mskTransDiscount.Enabled
            mskTransDiscount.SetFocus
        Case vbKeyL And CtrlDown = 4 And mskExtraCharges.Enabled
            mskExtraCharges.SetFocus
        Case vbKeyF And CtrlDown = 4 And mskTotalVAT.Enabled
            mskTotalVAT.SetFocus
        Case vbKeyF4 And cmdIndex(9).Enabled
            cmdIndex_Click 9
    End Select

End Function

Private Sub Form_Load()
    
    Dim lngRow As Long
    
    AddColumnsToGrid grdCommonTransactions, 44, GetSetting(strAppTitle, "Layout Strings", "grdCommonTransactions"), _
        "05NCNXCategoryID,04YCNCategoryShortDescription,04NCNXItemID,50YLNItemDescription,40NLNManufacturerDescription,10YRIQty,10YRFXUnitPrice,10NRFTotalNetPreDiscount,10YRFXDiscPercent,10YRFXDiscAmount,05YCNDiscAllow,10NRFTotalNetPostDiscount,10NRFXVATPercent,10NRFVATAmount,10" & IIf(txtRefersTo.text = "1", "N", "Y") & "RFTotalGross,10NRIXLastQty", _
        "ID Κατ,Κατ,ID Είδους,Είδος,Κατασκευαστής,Ποσότητα,Τιμή  μονάδος,Σύνολο,Ποσοστό έκπτωσης,Ποσό έκπτωσης,ΣΕ,Υπόλοιπο,Ποσοστό ΦΠΑ,Ποσό ΦΠΑ,Σύνολο,Τρέχουσα  ποσότητα"
    
    SetUpGrid lstIconList, grdCommonTransactions
    PositionControls Me, True, grdCommonTransactions: ColorizeControls Me, True
    
    ClearFields txtInvoiceTrnID, txtInvoicePersonID, txtInvoiceCodeID, txtInvoicePrintExtraRemarksID, txtInvoiceDeliveryPointID, txtInvoicePaymentWayID, txtInvoiceInDate, txtInvoiceInTime, txtInvoiceIsInvoiced, txtInvoiceIsPrinted, txtCodeDetailsID, txtCodeHandID, txtCodeLastNo, txtVATStateID, txtCodeInventoryQty, txtCodeInventoryValue, txtCodeTransformID, mskCodeLastDate, txtCodePrinterID, mskCodeDetailLines, txtProfession, txtAddress, txtCity, txtTaxNo, txtPhones, txtTaxOfficeDescription, grdCommonTransactions, txtCodeDateCheckID, txtCodeIsCreditID, txtCodeDTable, txtCodeDTableCredit
    ClearFields mskInvoiceIssueDate, txtPersonDescription, txtCodeShortDescription, lblCodeDescription, txtInvoiceNo, txtInvoicePrintExtraRemarks, txtDeliveryPointDescription, txtPaymentWayDescription, txtInvoicePlates, txtInvoiceRemarks, txtInvoiceTransportReason, txtInvoiceTransportWay, txtInvoiceLoadingSite, txtInvoiceDestinationSite, txtCodeShortDescriptionCredit, txtInvoiceNoCredit
    ClearFields mskTotalQty, mskTotalPreDiscount, mskDiscount, mskTransDiscount, mskTotalRestAmount, mskExtraCharges, mskTotalVAT, mskTotalGross
    
    DisableFields mskInvoiceIssueDate, txtPersonDescription, txtCodeShortDescription, txtInvoiceNo, txtInvoicePrintExtraRemarks, txtDeliveryPointDescription, txtPaymentWayDescription, txtInvoicePlates, txtInvoiceRemarks, txtInvoiceTransportReason, txtInvoiceTransportWay, txtInvoiceLoadingSite, txtInvoiceDestinationSite, txtCodeShortDescriptionCredit, txtInvoiceNoCredit
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8), cmdIndex(9)
    DisableFields mskDiscount, mskTransDiscount, mskTotalRestAmount, mskExtraCharges, mskTotalVAT
    
    HideCredits
    
    UpdateButtons Me, 5, 1, 0, 0, 1, 0, 1
    
    lngRow = 0

End Sub

Private Sub grdCommonTransactions_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
    
    Dim strCategoryID As String
    Dim strItemQuickDescription As String
    Dim strItemDescription As String
    
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
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
    
    Select Case lCol
        Case 2
            'Κατηγορία
            If grdCommonTransactions.CellValue(lRow, "CategoryShortDescription") <> "" Then
                Set tmpRecordset = CheckForMatch("CommonDB", grdCommonTransactions.CellValue(lRow, "CategoryShortDescription"), "Categories", "CategoryShortDescription", "String", 1, 1)
                tmpTableData = DisplayIndex(tmpRecordset, True, False, "Ευρετήριο", 3, 0, 1, 2, "ID", "Συντ.", "Περιγραφή", 0, 4, 40, 1, 1, 0)
                If tmpTableData.strCode = "" Then
                    FillCellWithSomething grdCommonTransactions, 0, grdCommonTransactions.CurRow, "6,7,8,9,10,11,13,14,15,16"
                End If
                If tmpTableData.strCode <> "" Then
                    grdCommonTransactions.CellValue(lRow, "CategoryID") = tmpTableData.strCode
                    grdCommonTransactions.CellValue(lRow, "CategoryShortDescription") = tmpTableData.strOneField
                    MoveToNextColumn grdCommonTransactions, lRow, lCol
                End If
            Else
                FillCellWithSomething grdCommonTransactions, "", grdCommonTransactions.CurRow, "1,2"
            End If
        Case 4
            'Είδος
            If grdCommonTransactions.CellValue(lRow, "ItemDescription") <> "" Then
                strCategoryID = IIf(grdCommonTransactions.CellValue(lRow, "CategoryID") <> "", "ItemCategoryID = " & grdCommonTransactions.CellValue(lRow, "CategoryID"), "")
                strItemQuickDescription = IIf(grdCommonTransactions.CellValue(lRow, "ItemDescription") <> "", grdCommonTransactions.CellValue(lRow, "ItemDescription"), "'")
                If Left(strItemQuickDescription, 1) <> "*" Then
                    strItemDescription = "Left(ItemQuickDescription, " & Len(strItemQuickDescription) & ") = '" & strItemQuickDescription & "'" & IIf(strCategoryID <> "", " AND " & strCategoryID, "")
                Else
                    If Len(strItemQuickDescription) > 1 Then
                        strItemDescription = "InStr(ItemQuickDescription, " & Right(strItemQuickDescription, Len(strItemQuickDescription) - 1) & ")" & IIf(strCategoryID <> "", " And " & strCategoryID, "")
                    Else
                        strItemDescription = strCategoryID
                    End If
                End If
                Set tmpRecordset = NewCheckForMatch("CommonDB", "ItemID, ItemCategoryID, ItemManufacturerID, CategoryDescription, ManufacturerDescription, ItemDescription, CategoryShortDescription, ItemVATPercent, ItemBalance, ItemActive, CategoryCheckBalance ", _
                    "((Items", _
                    "INNER JOIN Categories ON Items.ItemCategoryID = Categories.CategoryID) " & _
                    "INNER JOIN Manufacturers ON Items.ItemManufacturerID = Manufacturers.ManufacturerID) ", strItemDescription, "", "CategoryDescription, ManufacturerDescription, ItemDescription")
                tmpTableData = DisplayIndex(tmpRecordset, True, True, "Ευρετήριο", _
                    11, 0, 1, 2, 3, 5, 4, 6, 7, 8, 9, 10, _
                    "ID", "ID Κατηγορίας", "ID Κατασκευαστή", "Κατηγορία", "Περιγραφή", "Κατασκευαστής", "Συντ. κατηγορίας", "Φ.Π.Α.", "Υπόλοιπο", "", "Ε", _
                    0, 0, 0, 40, 50, 40, 0, 0, 10, 0, 0, _
                    0, 0, 0, 0, 0, 0, 0, 2, 2, 1, 1, "Items")
                If tmpTableData.strCode = "" Then
                    FillCellWithSomething grdCommonTransactions, "", grdCommonTransactions.CurRow, "3,4,5,11"
                    FillCellWithSomething grdCommonTransactions, "0", grdCommonTransactions.CurRow, "13,16"
                    ColorizeRowsWhenItemIsNotGiven lRow
                Else
                    grdCommonTransactions.CellValue(lRow, "ItemID") = tmpTableData.strCode
                    grdCommonTransactions.CellValue(lRow, "ItemDescription") = tmpTableData.strFourField
                    grdCommonTransactions.CellValue(lRow, "CategoryID") = tmpTableData.strOneField
                    grdCommonTransactions.CellValue(lRow, "CategoryShortDescription") = tmpTableData.strSixField
                    grdCommonTransactions.CellValue(lRow, "ManufacturerDescription") = tmpTableData.strFiveField
                    grdCommonTransactions.CellValue(lRow, "VATPercent") = IIf(txtVATStateID.text = "1", tmpTableData.strSevenField, "0")
                    grdCommonTransactions.CellValue(lRow, "LastQty") = tmpTableData.strEightField
                    ColorizeRowsWhenItemIsNotGiven lRow
                    MoveToNextColumn grdCommonTransactions, lRow, lCol
                End If
            Else
                FillCellWithSomething grdCommonTransactions, "", grdCommonTransactions.CurRow, "3,4,5,11"
                FillCellWithSomething grdCommonTransactions, "0", grdCommonTransactions.CurRow, "13,16"
                ColorizeRowsWhenItemIsNotGiven lRow
            End If
        Case 6
            'Ποσότητα
            If grdCommonTransactions.CellText(lRow, "Qty") <> "" Then MoveToNextColumn grdCommonTransactions, lRow, lCol
        Case 7
            'Τιμή μονάδος
            If grdCommonTransactions.CellText(lRow, "UnitPrice") <> "" Then MoveToNextColumn grdCommonTransactions, lRow, lCol
        Case 9
            'Ποσοστό έκπτωσης
            If Val(grdCommonTransactions.CellValue(lRow, "DiscPercent")) <> 0 Then
                If Val(grdCommonTransactions.CellValue(lRow, "DiscAmount")) = 0 Then
                    grdCommonTransactions.CellValue(lRow, "DiscAllow") = "Percent"
                End If
                MoveToNextColumn grdCommonTransactions, lRow, lCol + 2
            Else
                grdCommonTransactions.CellValue(lRow, "DiscAmount") = 0
                grdCommonTransactions.CellValue(lRow, "DiscAllow") = ""
                MoveToNextColumn grdCommonTransactions, lRow, lCol
            End If
        Case 10
            'Ποσό έκπτωσης
            If grdCommonTransactions.CellValue(lRow, "DiscAmount") <> 0 Then
                If Val(grdCommonTransactions.CellValue(lRow, "DiscPercent")) = 0 Then
                    grdCommonTransactions.CellValue(lRow, "DiscAllow") = "Amount"
                End If
                MoveToNextColumn grdCommonTransactions, lRow, lCol + 1
            Else
                grdCommonTransactions.CellValue(lRow, "DiscAmount") = 0
                grdCommonTransactions.CellValue(lRow, "DiscAllow") = ""
                MoveToNextColumn grdCommonTransactions, lRow, lCol + 1
            End If
        Case 15
            'Σύνολο
            DoReverseCalculation lRow
            MoveToNextColumn grdCommonTransactions, lRow, lCol
    End Select
    
    'If grdCommonTransactions.CellValue(lRow, "CategoryID") <> "" And grdCommonTransactions.CellValue(lRow, "ItemID") <> "" Then FillEmptyCellsWithZeros grdCommonTransactions, grdCommonTransactions.CurRow, "7,8,9,10,11,13,14,15,16"
    
    DoCalculations lRow
    CalculateTotals True
    
    blnGridEditInProgress = False
    
End Sub

Sub CalculateTotals(blnRecalculate As Boolean)

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
    
    On Error GoTo ErrTrap
    
    Dim lngRow As Long
    Dim lngCol As Long
    
    Dim intTotalQty As Integer
    
    Dim curTotalPreDiscount As Currency
    Dim curDiscount As Currency
    Dim curTotalTransDiscount As Currency
    Dim curTotalRestAmount As Currency
    Dim curExtraCharges As Currency
    Dim curTotalVAT As Currency
    Dim curTotalGross As Currency
    
    'Σύνολα
    For lngRow = 1 To grdCommonTransactions.RowCount
        'Ποσότητα
        intTotalQty = intTotalQty + grdCommonTransactions.CellValue(lngRow, "Qty")
        'Αξία προ έκπτωσης
        curTotalPreDiscount = curTotalPreDiscount + grdCommonTransactions.CellValue(lngRow, "TotalNetPreDiscount")
        'Έκπτωση
        curDiscount = curDiscount + grdCommonTransactions.CellValue(lngRow, "DiscAmount")
        'ΦΠΑ
        curTotalVAT = curTotalVAT + grdCommonTransactions.CellValue(lngRow, "VATAmount")
        'Γενικό σύνολο
        curTotalGross = curTotalGross + grdCommonTransactions.CellValue(lngRow, "TotalGross")
    Next lngRow
    
    'Υπόλοιπο αξίας
    curTotalRestAmount = Round(curTotalPreDiscount - curDiscount - CCur(mskTransDiscount.text), 2)
    'Λοιπές χρεώσεις
    curExtraCharges = CCur(mskExtraCharges.text)
    'ΦΠΑ
    curTotalVAT = IIf(blnRecalculate, IIf(mskTransDiscount.text <> "0,00" Or mskExtraCharges.text <> "0,00", (curTotalRestAmount + curExtraCharges) * (curExtraChargesVATPercent / 100), curTotalVAT), CCur(mskTotalVAT.text))
    'Γενικό σύνολο
    curTotalGross = curTotalRestAmount + curExtraCharges + curTotalVAT
    
    'Εμφανίζω
    mskTotalQty.text = format(intTotalQty, "#,##0")
    mskTotalPreDiscount.text = format(curTotalPreDiscount, "#,##0.00")
    mskDiscount.text = format(curDiscount, "#,##0.00")
    mskTotalVAT.text = format(curTotalVAT, "#,##0.00")
    mskTotalRestAmount.text = format(curTotalRestAmount, "#,##0.00")
    mskTotalGross.text = format(curTotalGross, "#,##0.00")
    
    'eΓραμμή
    CreateELine
        
    'Βγαίνω
    Exit Sub
    
ErrTrap:
    If Err.Number = 13 Then
        Resume Next
    End If

End Sub

Sub DoCalculations(lngRow As Long)

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
    
    On Error GoTo ErrTrap
    
    'Local μεταβλητές
    Dim curVATToAdd As Currency
    Dim curNetAmount As Currency
    Dim curDiscPerc As Currency
    Dim curDiscAmount As Currency
    Dim curRestAmount As Currency
    Dim curVATAmount As Currency
    Dim curGrossAmount As Currency
    
    'Αν είμαι σε μεταβολή και δεν έχω έκπτωση
    If Not blnStatus Then
        If grdCommonTransactions.CellValue(lngRow, "DiscPercent") = 0 And grdCommonTransactions.CellValue(lngRow, "DiscAmount") = 0 Then grdCommonTransactions.CellValue(lngRow, "DiscAllow") = ""
        If grdCommonTransactions.CellValue(lngRow, "DiscAmount") = 0 And grdCommonTransactions.CellValue(lngRow, "DiscPercent") = 0 Then grdCommonTransactions.CellValue(lngRow, "DiscAllow") = ""
    End If
    'Ποσότητα και Τιμή Μονάδας
    If grdCommonTransactions.CellValue(lngRow, "Qty") = "" Then grdCommonTransactions.CellValue(lngRow, "Qty") = 0
    If grdCommonTransactions.CellValue(lngRow, "UnitPrice") = "" Then grdCommonTransactions.CellValue(lngRow, "UnitPrice") = "0,00"
    'Υπολογισμός ποσότητας x τιμή μονάδος
    curNetAmount = grdCommonTransactions.CellValue(lngRow, "Qty") * grdCommonTransactions.CellValue(lngRow, "UnitPrice")
    'Υπολογισμός ποσοστού έκπτωσης αν έχω δώσει ποσό έκπτωσης
    If grdCommonTransactions.CellValue(lngRow, "DiscAllow") = "Amount" Then
        curDiscPerc = 100 * grdCommonTransactions.CellValue(lngRow, "DiscAmount") / curNetAmount
        curRestAmount = curNetAmount - grdCommonTransactions.CellValue(lngRow, "DiscAmount")
        curDiscAmount = grdCommonTransactions.CellValue(lngRow, "DiscAmount")
    End If
    'Υπολογισμός ποσού έκπτωσης αν έχω δώσει ποσοστό έκπτωσης
    If grdCommonTransactions.CellValue(lngRow, "DiscAllow") = "Percent" Then
        curDiscAmount = curNetAmount * (grdCommonTransactions.CellValue(lngRow, "DiscPercent") / 100)
        curRestAmount = curNetAmount - curDiscAmount
        curDiscPerc = grdCommonTransactions.CellValue(lngRow, "DiscPercent")
    End If
    'Υπολογισμός αξίας αν δεν έχω έκπτωση
    If grdCommonTransactions.CellValue(lngRow, "DiscAllow") = "" Or IsNull(grdCommonTransactions.CellValue(lngRow, "DiscAllow")) Then curRestAmount = curNetAmount - Val(grdCommonTransactions.CellValue(lngRow, "DiscAmount"))
    'Υπολογισμός ΦΠΑ
    curVATAmount = Round(curRestAmount * CCur(grdCommonTransactions.CellValue(lngRow, "VATPercent")) / 100, 2)
    'Υπολογισμός τελικής αξίας
    curGrossAmount = Round(curRestAmount, 2) + Round(curVATAmount, 2)
    
    'Στρογγυλοποίηση ΦΠΑ
    If ((txtRefersTo.text = "1" And blnRoundBuys = True) Or (txtRefersTo.text = "2" And blnRoundSales = True)) And Val(grdCommonTransactions.CellValue(lngRow, "DiscAllow")) = 0 And Val(grdCommonTransactions.CellValue(lngRow, "DiscPercent")) = 0 Then
        'Διόρθωση
        If Right(format(curGrossAmount, "#,##0.00"), 2) <= bytRoundCents And Right(format(curGrossAmount, "#,##0.00"), 2) <> "00" Then
            curVATAmount = curVATAmount - Right(CCur(curGrossAmount), 1) / 100
        End If
        'Προς τα πάνω
        If Right(format(curGrossAmount, "#,##0.00"), 2) >= 100 - bytRoundCents Then
            curVATToAdd = 1 - Right((curGrossAmount), 2) / 100
            curVATAmount = curVATAmount + curVATToAdd
        End If
    End If
    
    'Υπολογισμός τελικής αξίας
    curGrossAmount = Round(curRestAmount, 2) + Round(curVATAmount, 2)
    
    'Εμφανίζω
    grdCommonTransactions.CellValue(lngRow, "TotalNetPreDiscount") = curNetAmount
    grdCommonTransactions.CellValue(lngRow, "DiscPercent") = curDiscPerc
    grdCommonTransactions.CellValue(lngRow, "DiscAmount") = curDiscAmount
    grdCommonTransactions.CellValue(lngRow, "TotalNetPostDiscount") = curRestAmount
    grdCommonTransactions.CellValue(lngRow, "VATAmount") = curVATAmount
    grdCommonTransactions.CellValue(lngRow, "TotalGross") = curGrossAmount
    
    'Βγαίνω
    Exit Sub
    
ErrTrap:
    If Err.Number = 13 Or Err.Number = 6 Then Resume Next
    
End Sub

Sub DoReverseCalculation(lngRow As Long)

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
    
    On Error GoTo ErrTrap
    
    'Local μεταβλητές
    Dim strVAT As String
    Dim curNetAmount As Currency
    
    'Αποφορολόγηση
    strVAT = "1." & grdCommonTransactions.CellValue(lngRow, "VATPercent")
    'Καθαρή αξία
    curNetAmount = Replace(grdCommonTransactions.CellValue(lngRow, "TotalGross"), " ", "") / Val(strVAT)
    'Τιμή μονάδας
    grdCommonTransactions.CellValue(lngRow, "UnitPrice") = format(Round(curNetAmount / Val(grdCommonTransactions.CellValue(lngRow, "Qty")), 2), "#,##0.00")
    
    'Αν είναι μεγαλύτερη από 9999,99
    If grdCommonTransactions.CellValue(lngRow, "UnitPrice") > 9999.99 Then
        If MyMsgBox(4, lblTitle.Caption, Chr(13) & "Η τιμή μονάδας είναι υπερβολικά μεγάλη", 1) Then
        End If
        grdCommonTransactions.CellValue(lngRow, "UnitPrice") = "0,00"
    End If
    
    'Επαναυπολογίζω
    DoCalculations lngRow
    'Υπολογίζω τα σύνολα
    CalculateTotals True
    'Βγαίνω
    Exit Sub
    
ErrTrap:
    If Err.Number = 6 Or Err.Number = 11 Or Err.Number = 13 Then Resume Next

End Sub

Private Sub grdCommonTransactions_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid300_10Tec.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)

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
    
    If lCol = 7 Or lCol = 8 Or lCol = 9 Or lCol = 10 Or lCol = 15 Then
        vNewValue = Replace(sNewText, ".", ",")
        If vNewValue = "," Then
            vNewValue = "0,00"
        End If
    End If

End Sub

Private Sub grdCommonTransactions_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)

    Dim lngCol As Long
    Dim lngRow As Long
    Dim lngColCount As Long
    Dim lngRowCount As Long
    
    lngColCount = grdCommonTransactions.colCount
    lngRowCount = grdCommonTransactions.RowCount
    
    If grdCommonTransactions.RowCount = 0 Or grdCommonTransactions.CurRow = 0 Then Exit Sub
    
    grdCommonTransactions.Redraw = False
    
    For lngRow = 1 To lngRowCount
        For lngCol = 1 To lngColCount
            grdCommonTransactions.CellBackColor(lngRow, lngCol) = grdCommonTransactions.BackColor
        Next lngCol
    Next lngRow
    
    For lngCol = 1 To lngColCount
        grdCommonTransactions.CellBackColor(grdCommonTransactions.CurRow, lngCol) = RGB(128, 128, 128)
    Next lngCol
    
    grdCommonTransactions.Redraw = True

End Sub

Private Sub grdCommonTransactions_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdCommonTransactions_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    '1 = CategoryID
    '2 = CategoryShortDescription
    '3 = ItemID
    '4 = ItemDescription
    '5 = ManufacturerDescription
    
    Dim CtrlDown
    Dim lngRow As Long
    Dim lngCol As Long
    
    lngRow = grdCommonTransactions.CurRow
    lngCol = grdCommonTransactions.CurCol
    
    CtrlDown = Shift + vbCtrlMask
    
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    'F5 Πίνακας
    If KeyCode = vbKeyF5 Then
        'Κατηγορία
        If lngCol = 2 Then
            With UtilsItemCategories
                .Tag = "True"
                .Show 1, Me
            End With
        End If
        'Είδος
        If lngCol = 4 Then
            With Items
                .txtTable.text = "Items"
                .Tag = "True"
                .Show 1, Me
                If lngItemID <> 0 Then
                    Set tmpRecordset = NewCheckForMatch("CommonDB", "ItemID, ItemCategoryID, ItemManufacturerID, CategoryDescription, ManufacturerDescription, ItemDescription, CategoryShortDescription, ItemVATPercent", _
                    "((Items", _
                    "INNER JOIN Categories ON Items.ItemCategoryID = Categories.CategoryID) " & _
                    "INNER JOIN Manufacturers ON Items.ItemManufacturerID = Manufacturers.ManufacturerID) ", "ItemID = " & lngItemID, "", "CategoryDescription, ManufacturerDescription, ItemDescription")
                    tmpTableData = DisplayIndex(tmpRecordset, False, False, "Ευρετήριο", 8, 0, 1, 2, 3, 5, 4, 6, 7, "ID", "ID Κατηγορίας", "ID Κατασκευαστή", "Κατηγορία", "Περιγραφή", "Κατασκευαστής", "Συντ. κατηγορίας", "Φ.Π.Α.", 0, 0, 0, 40, 50, 40, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0)
                    If tmpTableData.strCode <> "" Then
                        grdCommonTransactions.CellValue(lngRow, "ItemID") = tmpTableData.strCode
                        grdCommonTransactions.CellValue(lngRow, "ItemDescription") = tmpTableData.strFourField
                        grdCommonTransactions.CellValue(lngRow, "CategoryID") = tmpTableData.strOneField
                        grdCommonTransactions.CellValue(lngRow, "CategoryShortDescription") = tmpTableData.strSixField
                        grdCommonTransactions.CellValue(lngRow, "ManufacturerDescription") = tmpTableData.strFiveField
                        grdCommonTransactions.CellValue(lngRow, "VATPercent") = IIf(txtVATStateID.text = "1", tmpTableData.strSevenField, "0")
                        ColorizeRowsWhenItemIsNotGiven lngRow
                        MoveToNextColumn grdCommonTransactions, lngRow, lngCol
                    End If
                End If
            End With
        End If
    End If
    
    'Πάνω βελάκι
    If KeyCode = 38 Then
        If grdCommonTransactions.CurRow = 1 Then
            grdCommonTransactions.CurCol = 0
            If txtInvoiceRemarks.Enabled Then txtInvoiceRemarks.SetFocus
            Exit Sub
        End If
    End If
    
    'Διαγραφή περιεχόμενου γραμμής CTRL + DEL
    If KeyCode = 46 And CtrlDown = 4 Then
        FillCellWithSomething grdCommonTransactions, "", grdCommonTransactions.CurRow, "1,2,3,4,5,11"
        FillCellWithSomething grdCommonTransactions, "0", grdCommonTransactions.CurRow, "6,7,8,9,10,12,13,14,15,16"
        ColorizeRowsWhenItemIsNotGiven grdCommonTransactions.CurRow
        grdCommonTransactions.SetCurCell grdCommonTransactions.CurRow, 2
        CalculateTotals True
    End If
    
    'Διαγραφή γραμμής CTRL + SHIFT + DEL
    If KeyCode = 46 And CtrlDown = 5 Then
        grdCommonTransactions.RemoveRow grdCommonTransactions.CurRow
        CalculateTotals True
    End If

End Sub

Private Sub grdCommonTransactions_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid300_10Tec.ETextEditFlags)

    blnGridEditInProgress = True
    
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
    
    'Απαγορεύεται πάντα
    If lCol = 1 Or lCol = 3 Or lCol = 5 Or lCol = 8 Or lCol = 11 Or lCol = 12 Or lCol = 13 Or lCol = 14 Then bCancel = True: Exit Sub
    
    'Απαγορεύεται, αν δεν έχω δώσει κατηγορία ή είδος
    If lCol >= 6 And (grdCommonTransactions.CellText(lRow, "CategoryID") = "" Or grdCommonTransactions.CellText(lRow, "ItemID") = "") Then bCancel = True: Exit Sub
    
    'Απαγορεύεται αν είμαι στο ποσοστό της έκπτωσης και έχω δώσει ποσό
    If lCol = 9 And grdCommonTransactions.CellText(lRow, "DiscAllow") = "Amount" Then bCancel = True: Exit Sub
    
    'Απαγορεύεται αν είμαι στο ποσό της έκπτωσης και έχω δώσει ποσοστό
    If lCol = 10 And grdCommonTransactions.CellText(lRow, "DiscAllow") = "Percent" Then bCancel = True: Exit Sub
    
    'Σύνολο - Αγορές απαγορεύεται - Πωλήσεις επιτρέπεται
    If txtRefersTo.text = "1" And lCol = 15 Then bCancel = True
    
    'Ποσότητα - Επιτρέπεται πάντα
    'If txtRefersTo.text = "1" And lCol = 6 Then bCancel = True
    
    'Σύνολο - Πωλήσεις επιτρέπεται αν είμαι σε παραστατικό που μεταβάλει τις αξίες
    If txtRefersTo.text = "2" And lCol = 15 And txtCodeInventoryValue.text = "" Then bCancel = True
    
    'Τιμή μονάδος σε παραστατικό που δεν μεταβάλεται
    If (lCol = 7 Or lCol = 9 Or lCol = 10) And txtCodeInventoryValue.text = "" Then bCancel = True
    
End Sub

Private Sub grdCommonTransactions_TextEditKeyPress(ByVal lRow As Long, ByVal lCol As Long, KeyAscii As Integer)

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

    'Ποσότητα
    If lCol = 6 Then
        If CheckForAcceptableKey(KeyAscii) Then
            CaptureNumbers grdCommonTransactions.TextEditText, lRow, lCol, KeyAscii, False
        Else
            KeyAscii = 0
        End If
    End If
    
    'Τιμή μονάδος, Ποσοστό έκπτωσης, Ποσό έκπτωσης, Σύνολο
    If lCol = 7 Or lCol = 9 Or lCol = 10 Or lCol = 11 Or lCol = 15 Then
        If CheckForAcceptableKey(KeyAscii) Then
            CaptureNumbers grdCommonTransactions.TextEditText, lRow, lCol, KeyAscii, True
        Else
            KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdCommonTransactions", grdCommonTransactions.LayoutCol

End Sub

Private Function ValidateFields()

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
    
    Dim lngRow As Long
    Dim lngCol As Long
    Dim intGivenColumns As Integer
    Dim intGivenRows As Integer
    
    ValidateFields = False
    
    'Ημερομηνία
    If Not CheckDateWithinLimits(strAppTitle, mskInvoiceIssueDate.text, datClosedPeriod) Then
        mskInvoiceIssueDate.SetFocus
        Exit Function
    End If
    
    'Συναλλασόμενος
    If DisplayMessage(1, 4, 1, "", txtInvoicePersonID.text) Then
        txtPersonDescription.SetFocus
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
        If DisplayMessage(2, 4, 1, "", "") Then
            txtInvoiceNo.SetFocus
            Exit Function
        End If
    End If
    
    'Για μηχανογραφικό παραστατικό σε νέα εγγραφή και με έλεγχο ημερομηνίας, η ημερομηνία πρέπει να είναι ίση με τη σημερινή
    If txtCodeHandID.text = "0" And blnStatus And txtCodeDateCheckID.text = "1" And CDate(mskInvoiceIssueDate.text) <> Date Then
        DisplayMessage 2, 4, 1, "", ""
        mskInvoiceIssueDate.SetFocus
        Exit Function
    End If
    
    'Για μηχανογραφικό παραστατικό σε νέα εγγραφή και με έλεγχο ημερομηνίας, το έτος της εγγραφής πρέπει να είναι το ίδιο με του τελευταίου παραστατικού
    If txtCodeHandID.text = "0" And blnStatus And txtCodeDateCheckID.text = "1" And Year(mskInvoiceIssueDate.text) <> Year(mskCodeLastDate.text) Then
        DisplayMessage 55, 4, 1, "", ""
        mskInvoiceIssueDate.SetFocus
        Exit Function
    End If
    
    'Στοιχείο ήδη καταχωρημένο: Ίδια ημερομηνία, ίδιος συναλλασόμενος, ίδιος τύπος παραστατικού, ίδιο νο παραστατικού
    If CheckForInvoiceExist(blnStatus, mskInvoiceIssueDate.text, txtInvoicePersonID.text, txtInvoiceCodeID.text, txtInvoiceNo.text) Then
        DisplayMessage 64, 4, 1, "", ""
        mskInvoiceIssueDate.SetFocus
        Exit Function
    End If
    
    'Νέα εγγραφή μηχανογραφικών πωλήσεων: Σωστή αυτόματη αρίθμηση
    If blnStatus And txtCodeHandID.text = "0" And txtRefersTo.text = "2" Then
        If CheckForValidSalesInvoiceNo Then
            DisplayMessage 55, 4, 1, "", ""
            mskInvoiceIssueDate.SetFocus
            Exit Function
        End If
    End If
    
    'Αν εκδίδεται ακυρωτικό, ελέγχω αν υπάρχει αυτό που ακυρώνω, όπως γίνεται και επάνω στα πεδία τη στιγμή της καταχώρησης
    If txtCodeIsCreditID.text = "1" Then
        If Not SeekInvoiceToCancel Then
            Exit Function
        End If
    End If
    
    'Εκτύπωση ΥγΣΜ - μόνο για πωλήσεις
    If txtInvoicePrintExtraRemarksID.text = "" And txtRefersTo.text = "2" Then
        If DisplayMessage(1, 4, 1, "", txtInvoicePrintExtraRemarksID.text) Then
            txtInvoicePrintExtraRemarks.SetFocus
            Exit Function
        End If
    End If
    
    'Τόπος παραλαβής - μόνο για αγορές
    If txtRefersTo.text = "1" Then
        If DisplayMessage(1, 4, 1, "", txtInvoiceDeliveryPointID.text) Then
            txtDeliveryPointDescription.SetFocus
            Exit Function
        End If
    End If
    
    'Τρόπος πληρωμής
    If DisplayMessage(1, 4, 1, "", txtInvoicePaymentWayID.text) Then
        txtPaymentWayDescription.SetFocus
        Exit Function
    End If
    
    'Είδη
    intGivenRows = 0
    grdCommonTransactions.CancelEdit
    
    For lngRow = 1 To grdCommonTransactions.RowCount
        
        intGivenColumns = 0
        
        For lngCol = 1 To 5
            If grdCommonTransactions.CellText(lngRow, lngCol) <> "" Then
                intGivenColumns = intGivenColumns + 1
            End If
        Next lngCol
        
        If intGivenColumns <> 0 Then
            If intGivenColumns <> 5 Then
                DisplayMessage 11, 4, 1, lngRow & " δεν είναι σωστή", ""
                grdCommonTransactions.SetFocus
                grdCommonTransactions.SetCurCell lngRow, "CategoryShortDescription"
                Exit Function
            Else
                intGivenRows = intGivenRows + 1
            End If
        End If
        
    Next lngRow
            
    If intGivenRows = 0 Then
        DisplayMessage 37, 4, 1, ""
        grdCommonTransactions.SetFocus
        grdCommonTransactions.SetCurCell 1, "CategoryShortDescription"
        Exit Function
    End If
    
    'MsgBox "Το παραστατικό θα αποθηκευτεί.", vbInformation
    
    ValidateFields = True

End Function

Private Sub mskExtraCharges_Validate(Cancel As Boolean)

    CalculateTotals True

End Sub

Private Sub mskInvoiceIssueDate_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 38 Then
        If grdCommonTransactions.TabStop = True Then
            grdCommonTransactions.SetFocus
            grdCommonTransactions.SetCurCell 1, 2
        End If
    End If
End Sub

Private Sub mskInvoiceIssueDate_Validate(Cancel As Boolean)

    grdCommonTransactions.Editable = CheckToEnableGrid
    grdCommonTransactions.TabStop = CheckToEnableGrid
    
    UpdateColTags
    
End Sub

Private Sub mskTotalVAT_Validate(Cancel As Boolean)

    CalculateTotals False
    
End Sub

Private Sub mskTransDiscount_Validate(Cancel As Boolean)

    CalculateTotals True

End Sub

Private Sub txtCodeShortDescription_Change()

    If txtCodeShortDescription.text = "" Then
        ClearFields txtInvoiceCodeID, lblCodeDescription, txtCodeDetailsID, txtCodeHandID, txtCodeLastNo, txtInvoiceNo, txtCodeInventoryQty, txtCodeInventoryValue, txtCodeTransformID, mskCodeLastDate, txtCodePrinterID, mskCodeDetailLines, txtInvoiceTransportReason, txtInvoiceLoadingSite, txtInvoiceTransportWay, txtInvoiceDestinationSite, txtCodeDateCheckID, txtCodeIsCreditID, txtCodeShortDescriptionCredit, txtInvoiceNoCredit, txtCodeDTable
        DisableFields txtInvoiceTransportReason, txtInvoiceLoadingSite, txtInvoiceTransportWay, txtInvoiceDestinationSite
        HideCredits
    End If
    
    CreateELine

End Sub

Private Sub txtCodeShortDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2
    If KeyCode = vbKeyF5 Then cmdIndex_Click 3

End Sub

Private Sub txtCodeShortDescription_Validate(Cancel As Boolean)

    If txtInvoiceCodeID.text = "" And txtCodeShortDescription.text <> "" Then
        cmdIndex_Click 2
        If txtInvoiceCodeID.text = "" Then Cancel = True
    End If
    
    grdCommonTransactions.Editable = CheckToEnableGrid
    grdCommonTransactions.TabStop = CheckToEnableGrid

    UpdateColTags
    CreateELine

End Sub

Private Sub txtCodeShortDescriptionCredit_Change()

    If txtCodeShortDescriptionCredit.text = "" Then
        ClearFields txtInvoiceCodeCreditID, txtCodeDTableCredit, txtInvoiceNoCredit, txtInvoiceCancelTrnID
    End If
    
    CreateELine

End Sub

Private Sub txtCodeShortDescriptionCredit_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 10
    
End Sub


Private Sub txtCodeShortDescriptionCredit_Validate(Cancel As Boolean)

    If txtInvoiceCodeCreditID.text = "" And txtCodeShortDescriptionCredit.text <> "" Then
        cmdIndex_Click 10
        If txtInvoiceCodeCreditID.text = "" Then Cancel = True
        If txtInvoiceCodeCreditID.text <> "" And txtInvoiceNoCredit.text <> "" Then
            SeekInvoiceToCancel
        End If
    End If
    
    CreateELine

End Sub

Private Sub txtDeliveryPointDescription_Change()

    If txtDeliveryPointDescription.text = "" Then ClearFields txtInvoiceDeliveryPointID

End Sub

Private Sub txtDeliveryPointDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 5
    If KeyCode = vbKeyF5 Then cmdIndex_Click 6

End Sub

Private Sub txtDeliveryPointDescription_Validate(Cancel As Boolean)

    If txtInvoiceDeliveryPointID.text = "" And txtDeliveryPointDescription.text <> "" Then cmdIndex_Click 5: If txtInvoiceDeliveryPointID.text = "" Then Cancel = True

End Sub

Private Sub txtInvoiceIsPrinted_Change()

    cmdIndex(9).Enabled = txtInvoicePersonID.text <> "" And txtInvoiceIsPrinted.text <> "1"

End Sub

Private Sub txtInvoiceNo_Validate(Cancel As Boolean)

    CreateELine
    
End Sub


Private Sub txtInvoiceNoCredit_Validate(Cancel As Boolean)

    If txtInvoiceCodeCreditID.text <> "" And txtInvoiceNoCredit.text <> "" Then
        SeekInvoiceToCancel
    End If
    
    CreateELine

End Sub



Private Sub txtInvoicePersonID_Change()

    cmdIndex(9).Enabled = txtInvoicePersonID.text <> "" And txtInvoiceIsPrinted.text <> "1"

End Sub

Private Sub txtInvoicePrintExtraRemarks_Change()

    If txtInvoicePrintExtraRemarks.text = "" And txtRefersTo.text = "2" Then ClearFields txtInvoicePrintExtraRemarksID
    
End Sub

Private Sub txtInvoicePrintExtraRemarks_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 4
    
End Sub


Private Sub txtInvoicePrintExtraRemarks_Validate(Cancel As Boolean)

    If txtInvoicePrintExtraRemarksID.text = "" And txtInvoicePrintExtraRemarks.text <> "" Then cmdIndex_Click 4: If txtInvoicePrintExtraRemarksID.text = "" Then Cancel = True
    
End Sub

Private Sub txtInvoiceRemarks_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Or KeyCode = 40 Then
        If grdCommonTransactions.TabStop = True Then
            grdCommonTransactions.SetCurCell 1, "CategoryShortDescription"
        End If
    End If

End Sub

Private Sub txtPaymentWayDescription_Change()

    If txtPaymentWayDescription.text = "" Then ClearFields txtInvoicePaymentWayID

End Sub

Private Sub txtPaymentWayDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 7
    If KeyCode = vbKeyF5 Then cmdIndex_Click 8

End Sub

Private Sub txtPaymentWayDescription_Validate(Cancel As Boolean)

    If txtInvoicePaymentWayID.text = "" And txtPaymentWayDescription.text <> "" Then cmdIndex_Click 7: If txtInvoicePaymentWayID.text = "" Then Cancel = True

End Sub

Private Sub txtPersonDescription_Change()

    If txtPersonDescription.text = "" Then
        ClearFields txtInvoicePersonID, txtInvoicePlates, txtProfession, txtAddress, txtCity, txtTaxNo, txtPhones, txtTaxOfficeDescription, txtVATStateID, txtInvoiceCodeID, txtCodeShortDescription, txtInvoiceNo, txtCodeShortDescriptionCredit, txtInvoiceNoCredit
    End If
    
    CreateELine

End Sub

Private Sub txtPersonDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0
    If KeyCode = vbKeyF5 Then cmdIndex_Click 1

End Sub

Private Sub txtPersonDescription_Validate(Cancel As Boolean)

    If txtInvoicePersonID.text = "" And txtPersonDescription.text <> "" Then
        cmdIndex_Click 0
        If txtInvoicePersonID.text = "" Then Cancel = True
    End If

    grdCommonTransactions.Editable = CheckToEnableGrid
    grdCommonTransactions.TabStop = CheckToEnableGrid
    
    UpdateColTags
    CreateELine

End Sub

Sub ShowCategoryTable(lngRow, lngCol)

    Dim tmpTableData As typTableData
    
    'tmpTableData = NewCheckForMatch("CommonDB", grdCommonTransactions.CellValue(lngRow, lngCol), "Categories", "CategoryShortDescription", "String", 0, 1, 1, 1, True, 5, 0, 1, 2, 3, 4, "", "ID", "Περιγραφή", "", "", 0, 5, 30, 0, 0, 0, 1, 0, 0, 0)
    'grdCommonTransactions.CellValue(lngRow, "CategoryID") = tmpTableData.strCode
    'grdCommonTransactions.CellValue(lngRow, "CategoryShortDescription") = tmpTableData.strOneField
    'grdCommonTransactions.CellValue(lngRow, "ItemDescriptionRequired") = tmpTableData.strThreeField
    'grdCommonTransactions.CellValue(lngRow, "CategoryCheckBalance") = tmpTableData.strFourField

End Sub

