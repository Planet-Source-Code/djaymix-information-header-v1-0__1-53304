VERSION 5.00
Begin VB.UserControl Button 
   AutoRedraw      =   -1  'True
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   BeginProperty Font 
      Name            =   "Wingdings 3"
      Size            =   8.25
      Charset         =   2
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   16
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   16
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   -480
      Top             =   -120
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Button
' unfinshed ********************

Option Explicit

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Dim Drawn           As Boolean
Dim CurPnt          As POINTAPI

Dim MOUSE_Hover     As Boolean
Dim MOUSE_Down      As Boolean

Dim WFP             As Long

Private Sub Timer1_Timer()
     GetCursorPos CurPnt
     
     WFP = WindowFromPoint(CurPnt.X, CurPnt.Y)
     
     If WFP = hWnd Then
          MOUSE_Hover = True
          
          Call Refresh
          
          If Timer1.Enabled = True Then Timer1.Enabled = False
     Else
          MOUSE_Hover = False
          
          Call Refresh
          
          If Timer1.Enabled = True Then Timer1.Enabled = False
     End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     MOUSE_Down = True
     
     Call Refresh
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Timer1.Enabled = False Then Timer1.Enabled = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     MOUSE_Down = False
     
     Call Refresh
End Sub

Sub Refresh()
     Cls
     'If Drawn = True Then Exit Sub
     
     Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), SystemColorConstants.vb3DShadow, B
     
     ForeColor = vbBlack
     CurrentX = (ScaleWidth - TextWidth("€")) \ 2
     CurrentY = (ScaleHeight - TextHeight("€")) \ 2
     
     If MOUSE_Hover Then
          If Not MOUSE_Down Then
               'CurrentX = CurrentX - 1
               'CurrentY = CurrentY - 1
          ElseIf MOUSE_Down Then
               CurrentX = CurrentX + 1
               CurrentY = CurrentY + 1
          End If
     End If
     
     Print "€"
     
     Drawn = True
End Sub

Private Sub UserControl_Resize()
     Call Refresh
End Sub
