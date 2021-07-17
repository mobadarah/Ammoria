VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      Height          =   4590
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7380
      Begin VB.PictureBox picLogo 
         Height          =   2385
         Left            =   510
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   2325
         ScaleWidth      =   1755
         TabIndex        =   1
         Top             =   855
         Width           =   1815
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "⁄„¯Ê—Ì«"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1125
         Left            =   3705
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Tag             =   "Product"
         Top             =   960
         Width           =   2280
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "»Ì∆… «·⁄„· : Windows XP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3915
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Tag             =   "Platform"
         Top             =   2400
         Width           =   3090
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·«’œ«— «·√Ê·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5700
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "Version"
         Top             =   2880
         Width           =   1305
      End
      Begin VB.Label lblWarning 
         Alignment       =   1  'Right Justify
         Caption         =   "Â–« «·»—‰«„Ã „Ã«‰Ì ·Ã„Ì⁄ «·⁄—»"
         Height          =   195
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Tag             =   "Warning"
         Top             =   3720
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
