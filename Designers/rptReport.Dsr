VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptReport 
   ClientHeight    =   11085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14985
   Icon            =   "rptReport.dsx":0000
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   26432
   _ExtentY        =   19553
   SectionData     =   "rptReport.dsx":1CCA
End
Attribute VB_Name = "rptReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_DataInitialize()

    Open strUnicodeFile For Input As #1
    
    Fields.RemoveAll
    
    Fields.Add "OneLongField"
    
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)

    Dim strLine As String
    Dim arr As String
    
    If VBA.EOF(1) Then
        EOF = True
        Exit Sub
    Else
        EOF = False
    End If
    
    Line Input #1, strLine
    
    Fields("OneLongField").Value = strLine
    
End Sub

Private Sub ActiveReport_ReportEnd()

    Close #1

End Sub

