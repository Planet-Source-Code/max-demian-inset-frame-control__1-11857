VERSION 5.00
Begin VB.UserControl InsetFrame 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   ScaleHeight     =   1215
   ScaleWidth      =   1515
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Inset Frame"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   840
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   1440
      X2              =   45
      Y1              =   1145
      Y2              =   1145
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   1440
      X2              =   1440
      Y1              =   1145
      Y2              =   165
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   45
      X2              =   45
      Y1              =   1145
      Y2              =   165
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   45
      X2              =   1440
      Y1              =   165
      Y2              =   165
   End
End
Attribute VB_Name = "InsetFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub SetSizes()
    Line1.X2 = (UserControl.Width - 50)
    Line4.X1 = (UserControl.Width - 50)
    Line4.Y1 = (UserControl.Height - 50)
    Line4.Y2 = (UserControl.Height - 50)

    Line2.Y1 = (UserControl.Height - 50)
    Line3.Y1 = (UserControl.Height - 50)
    Line3.X1 = (UserControl.Width - 50)
    Line3.X2 = (UserControl.Width - 50)

    If Label1.Caption = "" Then
        Label1.Visible = False
    Else
        Label1.Visible = True
    End If
End Sub

Private Sub UserControl_Initialize()
    Call SetSizes
End Sub

Private Sub UserControl_Resize()
    Call SetSizes
End Sub
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    Call SetSizes
    PropertyChanged "Caption"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Label1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Label1.Caption = PropBag.ReadProperty("Caption", "Inset Frame")
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000B)
    Label1.BackColor = PropBag.ReadProperty("BackColor", &H8000000B)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "Inset Frame")
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000B)
End Sub
