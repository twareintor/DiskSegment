VERSION 5.00
Object = "{7F0DC2FA-DACB-4A76-B3C3-86A36AB1228A}#1.0#0"; "LEDMeter.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DiskSegment"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDev 
      Caption         =   "..."
      Height          =   255
      Left            =   6360
      TabIndex        =   26
      Top             =   120
      Width           =   375
   End
   Begin LEDMETERLib.LEDMeter proHi 
      Height          =   255
      Left            =   240
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4560
      Width           =   6855
      _Version        =   65536
      _ExtentX        =   12091
      _ExtentY        =   450
      _StockProps     =   0
   End
   Begin VB.Frame frmSaveFile 
      Caption         =   "Save &Files"
      Height          =   1335
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   6855
      Begin VB.CommandButton cmdDir 
         Caption         =   "..."
         Height          =   255
         Left            =   6120
         TabIndex        =   25
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtMin 
         Height          =   285
         Left            =   1920
         TabIndex        =   16
         Text            =   "0"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtMax 
         Height          =   285
         Left            =   4560
         TabIndex        =   18
         Text            =   "0"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtOutputPath 
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label lblSizeBytes 
         Caption         =   "Size [Bytes]"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblMin 
         Caption         =   "mi&n"
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblMax 
         Caption         =   "&Max"
         Height          =   255
         Left            =   3960
         TabIndex        =   17
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblOutputPath 
         Caption         =   "Output &Path"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit..."
      Height          =   375
      Left            =   4680
      TabIndex        =   21
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   20
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "St&art"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   19
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Frame fraSegment 
      Caption         =   "Se&gment to process"
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   6855
      Begin VB.TextBox txtLowTo 
         Height          =   285
         Left            =   4800
         TabIndex        =   10
         Text            =   "0"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtLowFrom 
         Height          =   285
         Left            =   1920
         TabIndex        =   8
         Text            =   "0"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtHighTo 
         Height          =   285
         Left            =   4800
         TabIndex        =   6
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtHighFrom 
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblLowTo 
         Caption         =   "Low T&o"
         Height          =   255
         Left            =   3960
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblLowFrom 
         Caption         =   "Low F&rom"
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblHighTo 
         Caption         =   "High &To"
         Height          =   255
         Left            =   3960
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblHighFrom 
         Caption         =   "High &From"
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox txtDriveToProcess 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Text            =   "PhysicalDrive0"
      Top             =   120
      Width           =   4095
   End
   Begin LEDMETERLib.LEDMeter proLo 
      Height          =   255
      Left            =   240
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4920
      Width           =   6855
      _Version        =   65536
      _ExtentX        =   12091
      _ExtentY        =   450
      _StockProps     =   0
   End
   Begin VB.Label lblLog 
      Caption         =   "Click ""Start"" to start the program..."
      Height          =   735
      Left            =   240
      TabIndex        =   24
      Top             =   3720
      Width           =   6855
   End
   Begin VB.Label lblDriveToProcess 
      Caption         =   "&Drive To Read"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public bStop        As Boolean


Private Sub lblHigh_Click()

End Sub

Private Sub cmdDev_Click()
    frmDev.Show vbModal
    

End Sub

Private Sub cmdDir_Click()
    frmDir.Show vbModal
    frmDir.dirDir.Enabled = True
    txtOutputPath.Text = frmDir.dirDir.Path

End Sub

Private Sub cmdExit_Click()
    End

End Sub

Private Sub cmdStart_Click()
    cmdStop.Enabled = True
    cmdExit.Enabled = False
    modDiskSegment.StartRead_0010
    cmdExit.Enabled = True

End Sub

Private Sub cmdStop_Click()
    bStop = True

End Sub

Private Sub SetProgressBars()
    proHi.RedZone = 0
    proHi.YellowZone = 0
    proLo.RedZone = 0
    proHi.YellowZone = 0
    proHi.SetLevel 0
    proLo.SetLevel 0

End Sub

