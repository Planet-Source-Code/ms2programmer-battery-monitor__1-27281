VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Battery Monitor"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   735
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1635
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   0
      Width           =   735
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   480
         Top             =   600
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemPowerStatus Lib "kernel32" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long
Private Type SYSTEM_POWER_STATUS
        ACLineStatus As Byte
        BatteryFlag As Byte
        BatteryLifePercent As Byte
        Reserved1 As Byte
        BatteryLifeTime As Long
        BatteryFullLifeTime As Long
End Type
Dim Bat As SYSTEM_POWER_STATUS
Private Sub Fill_Box(Perc As Variant)
On Error Resume Next
Dim lPerc As Integer
    If lPerc = Perc Then Exit Sub
    For xXx = Picture1.Height To (Picture1.Height - (Picture1.Height * (Perc / 100))) Step -30
        Picture1.Line (0, xXx)-(Picture1.Width, xXx), RGB(Form1.Text1(0).Text, Form1.Text1(1).Text, Form1.Text1(1).Text)
    Next xXx
    lPerc = Perc
End Sub
Private Sub Form_Load()
    X = Picture1.Height - 50
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Form1.Visible = True
End Sub
Private Sub Timer1_Timer()
    Call GetSystemPowerStatus(Bat)
    Me.Caption = Bat.BatteryLifePercent & "%"
    Call Fill_Box(Bat.BatteryLifePercent)
    Timer1.Interval = 2000
End Sub
