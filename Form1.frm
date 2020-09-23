VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demian Net Inset Frame"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   360
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   6315
      TabIndex        =   2
      Top             =   720
      Width           =   6375
   End
   Begin Project1.InsetFrame InsetFrame2 
      Height          =   4575
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8070
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Demian Net Inset Frame"
      BackColor       =   255
   End
   Begin Project1.InsetFrame InsetFrame1 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Demian Net Inset Frame"
      BackColor       =   12632256
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Click()
    InsetFrame1.SetSizes
End Sub
