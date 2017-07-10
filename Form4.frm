VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8055
   LinkTopic       =   "Form4"
   ScaleHeight     =   3495
   ScaleWidth      =   8055
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton BlowPlus 
      Caption         =   "风量+"
      Height          =   735
      Left            =   6720
      TabIndex        =   22
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton BlowMinus 
      Caption         =   "风量-"
      Height          =   735
      Left            =   6720
      TabIndex        =   21
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton OFF 
      Caption         =   "OFF"
      Height          =   735
      Left            =   6720
      TabIndex        =   20
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton ExtCycle 
      Caption         =   "外循环"
      Height          =   735
      Left            =   5520
      TabIndex        =   19
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton InterCycle 
      Caption         =   "内循环"
      Height          =   735
      Left            =   4200
      TabIndex        =   18
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton MODE 
      Caption         =   "MODE"
      Height          =   735
      Left            =   2880
      TabIndex        =   17
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton DEF 
      Caption         =   "DEF"
      Height          =   735
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Display"
      Height          =   2175
      Left            =   1440
      TabIndex        =   3
      Top             =   240
      Width           =   5175
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   3600
         Top             =   840
      End
      Begin VB.TextBox TextDefrost 
         Height          =   495
         Left            =   3360
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox TextAC 
         Height          =   495
         Left            =   3360
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TextBlow 
         Height          =   495
         Left            =   3360
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TextMode 
         Height          =   495
         Left            =   960
         TabIndex        =   9
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox TextCycle 
         Height          =   495
         Left            =   960
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TextTemperature 
         Height          =   495
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "除雾"
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "AC开关"
         Height          =   375
         Left            =   2520
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "风量档位"
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "当前MODE"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "循环模式"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "设置温度"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton AC 
      Caption         =   "A/C"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton TempMinus 
      Caption         =   "温度-"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton TempPlus 
      Caption         =   "温度+"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IsTimeElaplse As Byte
Dim IsACActive As Byte
Dim IsDEFActive As Byte


Private Sub AC_Click()
IsACActive = Not IsACActive
Call ButtonSend(&H1, &H0)
Timer1.Enabled = False
Timer1.Enabled = True
Do
    If IsTimeElaplse Then
        IsTimeElaplse = 0
        Exit Do
    End If
    DoEvents
Loop
Call RequestResponse
If IsACActive Then
    AC.BackColor = RGB(0, 0, 255)
Else
    AC.BackColor = RGB(222, 222, 222)
End If
End Sub

Private Sub BlowMinus_Click()
Call ButtonSend(&H4, &H0)
Timer1.Enabled = False
Timer1.Enabled = True
Do
    If IsTimeElaplse Then
        IsTimeElaplse = 0
        Exit Do
    End If
    DoEvents
Loop
Call RequestResponse
End Sub

Private Sub BlowPlus_Click()
Call ButtonSend(&H8, &H0)
Timer1.Enabled = False
Timer1.Enabled = True
Do
    If IsTimeElaplse Then
        IsTimeElaplse = 0
        Exit Do
    End If
    DoEvents
Loop
Call RequestResponse
End Sub

Private Sub DEF_Click()
IsDEFActive = Not IsDEFActive
Call ButtonSend(&H10, &H0)
Timer1.Enabled = False
Timer1.Enabled = True
Do
    If IsTimeElaplse Then
        IsTimeElaplse = 0
        Exit Do
    End If
    DoEvents
Loop
Call RequestResponse

If IsDEFActive Then
    DEF.BackColor = RGB(0, 0, 255)
Else
    DEF.BackColor = RGB(222, 222, 222)
End If
End Sub

Private Sub ExtCycle_Click()
Call ButtonSend(&H40, &H0)
Timer1.Enabled = False
Timer1.Enabled = True
Do
    If IsTimeElaplse Then
        IsTimeElaplse = 0
        Exit Do
    End If
    DoEvents
Loop
Call RequestResponse
End Sub

Private Sub Form_Load()
IsTimeElaplse = 0
Timer1.Enabled = False

'获取之前设定的温度，模式，AC等作为初始值
Form1.Text1.Text = Form1.Text1.Text + vbCrLf + "C40D手动测试界面打开" + vbCrLf
Call RequestResponse
End Sub

Private Sub InterCycle_Click()
Call ButtonSend(&H20, &H0)
Timer1.Enabled = False
Timer1.Enabled = True
Do
    If IsTimeElaplse Then
        IsTimeElaplse = 0
        Exit Do
    End If
    DoEvents
Loop
Call RequestResponse
End Sub

Private Sub MODE_Click()
Call ButtonSend(&H0, &H2)
Timer1.Enabled = False
Timer1.Enabled = True
Do
    If IsTimeElaplse Then
        IsTimeElaplse = 0
        Exit Do
    End If
    DoEvents
Loop
Call RequestResponse
End Sub

Private Sub OFF_Click()
Call ButtonSend(&H0, &H4)
Timer1.Enabled = False
Timer1.Enabled = True
Do
    If IsTimeElaplse Then
        IsTimeElaplse = 0
        Exit Do
    End If
    DoEvents
Loop
Call RequestResponse
End Sub

Private Sub TempMinus_Click()
Call ButtonSend(&H80, &H0)
Timer1.Enabled = False
Timer1.Enabled = True
Do
    If IsTimeElaplse Then
        IsTimeElaplse = 0
        Exit Do
    End If
    DoEvents
Loop
Call RequestResponse
End Sub

Private Sub TempPlus_Click()
Call ButtonSend(&H0, &H1)
Timer1.Enabled = False
Timer1.Enabled = True
Do
    If IsTimeElaplse Then
        IsTimeElaplse = 0
        Exit Do
    End If
    DoEvents
Loop
Call RequestResponse
End Sub

Private Sub ButtonSend(data0 As Byte, data1 As Byte)
Dim i As Byte
t_frame(0) = t_head(0)
t_frame(1) = t_head(1)
t_frame(2) = 17
t_frame(3) = &HF0
t_frame(4) = &H1
t_frame(5) = &H5    'ID
t_frame(6) = &H12
t_frame(7) = 8      'DLC
t_frame(8) = 0      'CYCLE
t_frame(9) = data0    'DATA 9-16
t_frame(10) = data1
t_frame(11) = &H10
t_frame(12) = &H1
t_frame(13) = &H0
t_frame(14) = &H0
t_frame(15) = &H0
t_frame(16) = &H0
t_checksum = 0      '将上一次的累加和清零
For i = 2 To 16
t_checksum = t_checksum + CLng(t_frame(i))
Next i
t_frame(17) = t_checksum \ 256
t_frame(18) = t_checksum Mod 256
t_frame(19) = t_tail(0)
t_frame(20) = t_tail(1)

Form1.MSComm1.Output = t_frame
End Sub


Private Sub RequestResponse()
t_frame(0) = t_head(0)
t_frame(1) = t_head(1)
t_frame(2) = &H7
t_frame(3) = &HF0
t_frame(4) = &H7
t_frame(5) = &H5    'ID &H0513
t_frame(6) = &H13
t_checksum = CLng(t_frame(2)) + t_frame(3) + t_frame(4) + t_frame(5) + t_frame(6)
t_frame(7) = t_checksum \ 256
t_frame(8) = t_checksum Mod 256
t_frame(9) = t_tail(0)
t_frame(10) = t_tail(1)

Form1.MSComm1.Output = t_frame
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
IsTimeElaplse = 1
End Sub
