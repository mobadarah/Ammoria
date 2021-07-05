VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Parsed Tree"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   10725
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   10455
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    On Error Resume Next
    Text1.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200

End Sub
