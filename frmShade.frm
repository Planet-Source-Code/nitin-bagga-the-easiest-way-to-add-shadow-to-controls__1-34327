VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   1605
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Shape FocusShadow 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   480
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub SamShadeMe(ctrl As Control, Shade As Integer)
On Error Resume Next:
    If Shade = True Then
        FocusShadow.Top = ctrl.Top + 70
        FocusShadow.Left = ctrl.Left + 70
        FocusShadow.Height = ctrl.Height
        FocusShadow.Width = ctrl.Width
        FocusShadow.Visible = True
    Else
        FocusShadow.Visible = False
    End If
End Sub

Private Sub Text2_GotFocus()
SamShadeMe Text2, True
End Sub

Private Sub Text1_GotFocus()
SamShadeMe Text1, True
End Sub

Private Sub Text3_GotFocus()
SamShadeMe Text3, True
End Sub

Private Sub Text1_LostFocus()
SamShadeMe Text1, True
End Sub

Private Sub Text2_LostFocus()
SamShadeMe Text2, True
End Sub

Private Sub Text3_LostFocus()
SamShadeMe Text3, True
End Sub
