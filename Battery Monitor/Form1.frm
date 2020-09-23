VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Battery Monitor"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Dockable Window"
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2040
      TabIndex        =   20
      Top             =   2760
      Width           =   1815
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "0"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0"
      Top             =   2280
      Width           =   495
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   14
      Top             =   2280
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Increment       =   5
      Max             =   255
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "0"
      Top             =   2280
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Battery Details:"
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   50
      Width           =   3735
      Begin VB.Label Label1 
         Caption         =   "Charging/Unplugged  State:"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label AnsL 
         Caption         =   "Label2"
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   10
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label AnsL 
         Caption         =   "Label2"
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   9
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label AnsL 
         Caption         =   "Label2"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   8
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label AnsL 
         Caption         =   "Label2"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label AnsL 
         Caption         =   "Label2"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Battery Full Lifetime:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Battery Lifetime:"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Battery Life Percent:"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Battery Flag:"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3495
      Left            =   3960
      ScaleHeight     =   3435
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   480
         Top             =   600
      End
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   15
      Top             =   2280
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Value           =   255
      Increment       =   5
      Max             =   255
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   17
      Top             =   2280
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Increment       =   5
      Max             =   255
      Enabled         =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "RGB Color Value:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2355
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
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
        Picture1.Line (0, xXx)-(Picture1.Width, xXx), RGB(Text1(0).Text, Text1(1).Text, Text1(1).Text)
    Next xXx
    lPerc = Perc
End Sub
Private Sub Command1_Click()
    Unload Me
    End
End Sub
Private Sub Command2_Click()
    Me.Visible = False
    Form2.Show
End Sub
Private Sub Form_Load()
On Error Resume Next
    UpDown1(0).Value = GetFromINI("RGB", "R", Left(App.Path, 3) & "BatMonitor.ini")
    UpDown1(1).Value = GetFromINI("RGB", "G", Left(App.Path, 3) & "BatMonitor.ini")
    UpDown1(2).Value = GetFromINI("RGB", "B", Left(App.Path, 3) & "BatMonitor.ini")
    X = Picture1.Height - 50
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call WriteToINI("RGB", "R", Text1(0).Text, Left(App.Path, 3) & "BatMonitor.ini")
    Call WriteToINI("RGB", "G", Text1(1).Text, Left(App.Path, 3) & "BatMonitor.ini")
    Call WriteToINI("RGB", "B", Text1(2).Text, Left(App.Path, 3) & "BatMonitor.ini")
    End
End Sub
Private Sub Timer1_Timer()
    Call GetSystemPowerStatus(Bat)
    Call Fill_Box(Bat.BatteryLifePercent)
    If Bat.BatteryLifePercent <> 255 Then ProgressBar1.Value = Bat.BatteryLifePercent
    Call Load_Data
    Timer1.Interval = 2000
End Sub
Private Sub Load_Data()
    AnsL(0).Caption = IIf(Bat.ACLineStatus = 1, "Plugged In", "Unplugged!")
    AnsL(1).Caption = IIf(Bat.BatteryFlag = 9, "Charging...", "Not Charging...")
    AnsL(2).Caption = Bat.BatteryLifePercent & "%"
    AnsL(3).Caption = IIf(Bat.BatteryLifeTime = -1, "Unkown!", Bat.BatteryLifeTime)
    AnsL(4).Caption = IIf(Bat.BatteryFullLifeTime = -1, "Unkown!", Bat.BatteryFullLifeTime)
    If Bat.BatteryLifePercent = 255 Then
        For q = 0 To 4
            AnsL(q).Caption = "Battery Removed!"
        Next q
    End If
End Sub
Private Sub UpDown1_Change(Index As Integer)
    Text1(Index).Text = UpDown1(Index).Value
    Call Fill_Box(Bat.BatteryLifePercent + 1)
End Sub
