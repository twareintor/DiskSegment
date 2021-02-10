Attribute VB_Name = "modMain"
Option Explicit



Dim Ini             As String
Dim Num             As String
Dim Nam             As String
Dim Des             As String
Dim Pat             As String
Dim Exe             As String
Dim Aut             As String
Dim C00             As String
Dim C01             As String
Dim C02             As String
    

Public Sub main()
    Initializations
    Load frmDisplay
    frmDisplay.txtLog.Text = "********************************************************************"
    frmDisplay.txtLog.Text = vbCrLf & frmDisplay.txtLog.Text & Num
    frmDisplay.txtLog.Text = vbCrLf & frmDisplay.txtLog.Text & Nam
    frmDisplay.txtLog.Text = vbCrLf & frmDisplay.txtLog.Text & Des
    frmDisplay.txtLog.Text = vbCrLf & frmDisplay.txtLog.Text & Pat
    frmDisplay.txtLog.Text = vbCrLf & frmDisplay.txtLog.Text & Exe
    frmDisplay.txtLog.Text = vbCrLf & frmDisplay.txtLog.Text & Aut
    frmDisplay.txtLog.Text = vbCrLf & frmDisplay.txtLog.Text & C00
    frmDisplay.txtLog.Text = vbCrLf & frmDisplay.txtLog.Text & C01
    frmDisplay.txtLog.Text = vbCrLf & frmDisplay.txtLog.Text & C02
    frmDisplay.Show vbModal
    frmMain.Show
    
End Sub

Private Sub Initializations()
    Ini = "Version Features: Nr. "
    Num = "1.0.0/2010-10-22 00:00"
    Nam = "Disk Segment - Data Recovery Utility"
    Des = "It scans the drive and search for lost data (JPEG headers)"
    Pat = "Project Application: """ & App.Path & """"
    Exe = "Program name: " & App.EXEName & ".exe"
    Aut = "Ciutacu Claudiu - info@kwatech.at"
    C00 = "This program is provides as-is with absolutely NO WARRANTY"
    C01 = "Press <ESC> key to start the program..."

End Sub
