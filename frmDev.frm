VERSION 5.00
Begin VB.Form frmDev 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select device to read..."
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.DriveListBox drvDev 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblDev 
      Caption         =   "&Device"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long



Private Sub cmdOK_Click()
    Me.Hide
    frmMain.txtDriveToProcess = DosDevice

End Sub

Private Function DosDevice() As String
    '//
    Dim strOut          As String       '// output value
    Dim intOut          As Long         '// another output
    Dim strDsk          As String
    Dim strDev          As String
    Dim intNCh          As Long
    Dim strTmp          As String       '// takes the result of the drvDev, in order to set it in a friendly format...
        
    strTmp = frmDev.drvDev.Drive
    strDsk = Left(strTmp, InStr(strTmp, ":"))
    strDev = "                                                                                   "
    intNCh = 255
    intOut = QueryDosDevice(strDsk, strDev, intNCh)
    Debug.Print strDev
    Debug.Print intOut
    strDev = Trim(strDev)
    strOut = strDev
    DosDevice = strOut
 
End Function


