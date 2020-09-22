VERSION 5.00
Object = "*\AGradientInfoHeader.vbp"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Information Header v1.0"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Show Example"
      Height          =   495
      Left            =   5055
      TabIndex        =   23
      Top             =   7350
      Width           =   1605
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customize"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5730
      Left            =   105
      TabIndex        =   1
      Top             =   1530
      Width           =   6585
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   4230
         TabIndex        =   22
         Text            =   "Combo3"
         Top             =   5115
         Width           =   1515
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   3180
         List            =   "Form1.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   870
         Width           =   2730
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Choose Font"
         Height          =   315
         Left            =   2655
         TabIndex        =   19
         Top             =   5115
         Width           =   1500
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Choose Forecolor"
         Height          =   315
         Left            =   2655
         TabIndex        =   18
         Top             =   4740
         Width           =   1500
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         Height          =   270
         Left            =   4710
         TabIndex        =   17
         Top             =   3765
         Width           =   510
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Multi Line"
         Height          =   195
         Left            =   2685
         TabIndex        =   16
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2685
         TabIndex        =   14
         Text            =   "Information Header v1.0|Created By: Jayson Ragasa|Copyright© 2004 Baguio City|Philippines"
         Top             =   3765
         Width           =   1995
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":0024
         Left            =   2700
         List            =   "Form1.frx":0034
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3030
         Width           =   2565
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Has Icon"
         Height          =   210
         Left            =   1245
         TabIndex        =   10
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse for picture or icon"
         Height          =   480
         Left            =   1260
         TabIndex        =   9
         Top             =   3060
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   585
         Max             =   100
         Min             =   1
         TabIndex        =   5
         Top             =   915
         Value           =   100
         Width           =   2430
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000010&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Index           =   0
         Left            =   405
         ScaleHeight     =   1020
         ScaleWidth      =   525
         TabIndex        =   3
         ToolTipText     =   "Double Click to choose color"
         Top             =   1245
         Width           =   585
      End
      Begin VB.PictureBox Picture1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Index           =   1
         Left            =   5535
         ScaleHeight     =   1020
         ScaleWidth      =   525
         TabIndex        =   2
         ToolTipText     =   "Double Click to choose color"
         Top             =   1245
         Width           =   585
      End
      Begin GradientInfoHeader.InfoHeader InfoHeader1 
         Height          =   1455
         Left            =   1245
         TabIndex        =   4
         Top             =   1245
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   2566
         GradientStyle   =   0
         LeftColor       =   8421504
         RightColor      =   13160660
         MaxFill         =   100
         FontName        =   "Tahoma"
         ForeColor       =   4210752
         Caption         =   "Information Header v1.0"
         MultiLine       =   0   'False
         Alignment       =   0
         HasIcon         =   -1  'True
         Picture         =   "Form1.frx":0059
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Gradient Style"
         Height          =   195
         Index           =   1
         Left            =   3195
         TabIndex        =   20
         Top             =   660
         Width           =   1500
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Info: if MultiLine is Checked, use pipe '|' as vbCrLf"
         Height          =   435
         Index           =   2
         Left            =   2670
         TabIndex        =   15
         Top             =   4275
         Width           =   2670
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Caption"
         Height          =   195
         Index           =   1
         Left            =   2715
         TabIndex        =   13
         Top             =   3510
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Alignment Style"
         Height          =   195
         Index           =   0
         Left            =   2715
         TabIndex        =   11
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Caption         =   "Click to change Right Color"
         ForeColor       =   &H80000017&
         Height          =   810
         Index           =   1
         Left            =   5565
         TabIndex        =   8
         Top             =   2370
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Caption         =   "Click to change Left Color"
         ForeColor       =   &H80000017&
         Height          =   810
         Index           =   0
         Left            =   405
         TabIndex        =   7
         Top             =   2370
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scroll to change Gradient Max Fill"
         Height          =   195
         Index           =   0
         Left            =   645
         TabIndex        =   6
         Top             =   660
         Width           =   2370
      End
   End
   Begin GradientInfoHeader.InfoHeader InfoHeader2 
      Height          =   1155
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   2037
      GradientStyle   =   1
      LeftColor       =   0
      RightColor      =   8388608
      MaxFill         =   100
      FontName        =   "Tahoma"
      ForeColor       =   -2147483639
      Caption         =   $"Form1.frx":03F3
      MultiLine       =   -1  'True
      Alignment       =   0
      HasIcon         =   -1  'True
      Picture         =   "Form1.frx":0439
   End
   Begin MSComDlg.CommonDialog CDlg 
      Left            =   6360
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
     InfoHeader1.HasIcon = CBool(Check1.Value)
     
     Command1.Enabled = CBool(Check1.Value)
End Sub

Private Sub Check2_Click()
     InfoHeader1.MultiLine = CBool(Check2.Value)
End Sub

Private Sub Combo1_Click()
     InfoHeader1.Alignment = Combo1.ListIndex
End Sub

Private Sub Combo2_Click()
     InfoHeader1.GradientStyle = Combo2.ListIndex
End Sub

Private Sub Combo3_Click()
     InfoHeader1.FontName = Combo3.List(Combo3.ListIndex)
End Sub

Private Sub Command1_Click()
     With CDlg
          .Filter = "Picture or Icon files|*.ico;*.bmp;*.gif;*.jpg"
          .DialogTitle = "Open"
          .FileName = vbNullString
          .ShowOpen
          
          If .FileName <> vbNullString Then
               Set InfoHeader1.Picture = LoadPicture(.FileName)
          End If
     End With
End Sub

Private Sub Command2_Click()
     InfoHeader1.Caption = Text1.Text
End Sub

Private Sub Command3_Click()
     CDlg.ShowColor
     
     InfoHeader1.ForeColor = CDlg.Color
End Sub

Private Sub Command4_Click()
     Dim i     As Integer
     
     Combo3.Clear
     
     For i = 0 To Screen.FontCount - 1
          Combo3.AddItem Screen.Fonts(i)
     Next i
End Sub

Private Sub Command5_Click()
     frmExample.Show 0, Me
End Sub

Private Sub HScroll1_Scroll()
     InfoHeader1.MaxFill = HScroll1.Value
End Sub

Private Sub Picture1_Click(Index As Integer)
     CDlg.ShowColor
     
     Picture1(Index).BackColor = CDlg.Color
     
     If Index = 0 Then
          InfoHeader1.LeftColor = CDlg.Color
     ElseIf Index = 1 Then
          InfoHeader1.RightColor = CDlg.Color
     End If
End Sub
