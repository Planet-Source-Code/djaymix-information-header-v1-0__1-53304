VERSION 5.00
Object = "*\AGradientInfoHeader.vbp"
Begin VB.Form frmExample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Example"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   6210
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5280
      Top             =   6435
   End
   Begin GradientInfoHeader.InfoHeader InfoHeader1 
      Height          =   390
      Left            =   105
      TabIndex        =   0
      Top             =   1545
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   688
      GradientStyle   =   0
      LeftColor       =   8421504
      RightColor      =   13160660
      MaxFill         =   100
      FontName        =   "Tahoma"
      ForeColor       =   -2147483639
      Caption         =   "Code Settings"
      MultiLine       =   0   'False
      Alignment       =   0
      HasIcon         =   -1  'True
      Picture         =   "frmExample.frx":0000
   End
   Begin GradientInfoHeader.InfoHeader InfoHeader2 
      Height          =   555
      Left            =   105
      TabIndex        =   1
      Top             =   3630
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   979
      GradientStyle   =   0
      LeftColor       =   13160660
      RightColor      =   16744576
      MaxFill         =   100
      FontName        =   "Tahoma"
      ForeColor       =   0
      Caption         =   "Editor Settings"
      MultiLine       =   0   'False
      Alignment       =   0
      HasIcon         =   -1  'True
      Picture         =   "frmExample.frx":0452
   End
   Begin GradientInfoHeader.InfoHeader InfoHeader3 
      Height          =   1275
      Left            =   105
      TabIndex        =   2
      Top             =   135
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2249
      GradientStyle   =   0
      LeftColor       =   16777215
      RightColor      =   13160660
      MaxFill         =   100
      FontName        =   "Tahoma"
      ForeColor       =   0
      Caption         =   "Example"
      MultiLine       =   0   'False
      Alignment       =   0
      HasIcon         =   -1  'True
      Picture         =   "frmExample.frx":5C44
   End
   Begin VB.Image Image2 
      Height          =   2865
      Left            =   150
      Picture         =   "frmExample.frx":691E
      Top             =   4200
      Width           =   5745
   End
   Begin VB.Image Image1 
      Height          =   1470
      Left            =   195
      Picture         =   "frmExample.frx":3C4E0
      Top             =   1980
      Width           =   5025
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
     Dim tmp As String
     
     tmp = "This is an example where you going to use|the Information Header Tool|Current Time: " & Time & "|Current Date: " & Date
     
     InfoHeader3.Caption = tmp
     InfoHeader3.MultiLine = True
End Sub

Private Sub Timer1_Timer()
     Dim tmp As String
     
     tmp = "This is an example where you going to use|the Information Header Tool|Current Time: " & Time & "|Current Date: " & Date
     
     InfoHeader3.Caption = tmp
     InfoHeader3.MultiLine = True
End Sub
