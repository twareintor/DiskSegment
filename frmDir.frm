VERSION 5.00
Begin VB.Form frmDir 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Directory"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   3240
      Width           =   1695
   End
   Begin VB.DirListBox dirDir 
      Height          =   2340
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   6135
   End
   Begin VB.DriveListBox drvDir 
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lblDir 
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    Me.Hide

End Sub

Private Sub cmdOK_Click()
    Me.Hide

End Sub

Private Sub drvDir_Change()
    Me.dirDir.Path = Me.drvDir.Drive

End Sub
