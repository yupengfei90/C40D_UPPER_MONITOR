VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "AD采集界面"
   ClientHeight    =   8775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10155
   LinkTopic       =   "Form2"
   ScaleHeight     =   8775
   ScaleWidth      =   10155
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame8 
      Caption         =   "Andata7"
      Height          =   4095
      Left            =   7800
      TabIndex        =   63
      Top             =   4440
      Width           =   2055
      Begin VB.CommandButton SIG_Sensor_AD0 
         Caption         =   "SIG_Sensor_AD0"
         Height          =   375
         Left            =   120
         TabIndex        =   71
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Sensor_AD4 
         Caption         =   "SIG_Sensor_AD4"
         Height          =   375
         Left            =   120
         TabIndex        =   70
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton Reserved_AD0701 
         Caption         =   "Reserved_AD0701"
         Height          =   375
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Trig_Low_AD7 
         Caption         =   "SIG_Trig_Low_AD7"
         Height          =   375
         Left            =   120
         TabIndex        =   68
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Reserved_AD0700 
         Caption         =   "Reserved_AD0700"
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton Reserved_AD0703 
         Caption         =   "Reserved_AD0703"
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Sensor_AD5 
         Caption         =   "SIG_Sensor_AD5"
         Height          =   375
         Left            =   120
         TabIndex        =   65
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Trig_Low_AD8 
         Caption         =   "SIG_Trig_Low_AD8"
         Height          =   375
         Left            =   120
         TabIndex        =   64
         Top             =   3120
         Width           =   1695
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Andata6"
      Height          =   4095
      Left            =   7800
      TabIndex        =   54
      Top             =   120
      Width           =   2055
      Begin VB.CommandButton SIG_Sensor_AD3 
         Caption         =   "SIG_Sensor_AD3"
         Height          =   375
         Left            =   120
         TabIndex        =   62
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Sensor_AD7 
         Caption         =   "SIG_Sensor_AD7"
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Sensor_AD2 
         Caption         =   "SIG_Sensor_AD2"
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Trig_Low_AD5 
         Caption         =   "SIG_Trig_Low_AD5"
         Height          =   375
         Left            =   120
         TabIndex        =   59
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Sensor_AD1 
         Caption         =   "SIG_Sensor_AD1"
         Height          =   375
         Left            =   120
         TabIndex        =   58
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Trig_Low_AD6 
         Caption         =   "SIG_Trig_Low_AD6"
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Trig_Low_AD4 
         Caption         =   "SIG_Trig_Low_AD4"
         Height          =   375
         Left            =   120
         TabIndex        =   56
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Sensor_AD6 
         Caption         =   "SIG_Sensor_AD6"
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   3120
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Andata5"
      Height          =   4095
      Left            =   5280
      TabIndex        =   45
      Top             =   4440
      Width           =   2055
      Begin VB.CommandButton Location_Step_AD3 
         Caption         =   "Location_Step_AD3"
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CommandButton KL15_Cur_AD 
         Caption         =   "KL15_Cur_AD"
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton KL30_Cur_AD 
         Caption         =   "KL30_Cur_AD"
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton POWER_Cur_AD 
         Caption         =   "POWER_Cur_AD"
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton KL15_Vol_AD 
         Caption         =   "KL15_Vol_AD"
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton Reserved_AD0003 
         Caption         =   "Reserved_AD0003"
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton Location_Step_AD4 
         Caption         =   "Location_Step_AD4"
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton Location_DC_motor_AD4 
         Caption         =   "Location_DC_motor_AD4"
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   3120
         Width           =   1695
      End
   End
   Begin VB.CommandButton SIG_Trig_Low_AD2 
      Caption         =   "SIG_Trig_Low_AD2"
      Height          =   375
      Left            =   5400
      TabIndex        =   44
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Frame Frame5 
      Caption         =   "Andata4"
      Height          =   4095
      Left            =   5280
      TabIndex        =   36
      Top             =   120
      Width           =   2055
      Begin VB.CommandButton SIG_Trig_Low_AD3 
         Caption         =   "SIG_Trig_Low_AD3"
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Trig_Hig_AD0 
         Caption         =   "SIG_Trig_Hig_AD0"
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton KL30_Vol_AD 
         Caption         =   "KL30_Vol_AD"
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton Power_Vol_AD 
         Caption         =   "Power_Vol_AD"
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Trig_Hig_AD1 
         Caption         =   "SIG_Trig_Hig_AD1"
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Trig_Hig_AD2 
         Caption         =   "SIG_Trig_Hig_AD2"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Trig_Hig_AD3 
         Caption         =   "SIG_Trig_Hig_AD3"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   2160
         Width           =   1695
      End
   End
   Begin VB.CommandButton Motor_Step_Cur_AD 
      Caption         =   "Motor_Step_Cur_AD"
      Height          =   375
      Left            =   2880
      TabIndex        =   35
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton SIG_Trig_Low_Cur_AD 
      Caption         =   "SIG_Trig_Low_Cur_AD"
      Height          =   375
      Left            =   2880
      TabIndex        =   34
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton SIG_Trig_Hig_Cur_AD 
      Caption         =   "SIG_Trig_Hig_Cur_AD"
      Height          =   375
      Left            =   2880
      TabIndex        =   33
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Caption         =   "Andata3"
      Height          =   4095
      Left            =   2760
      TabIndex        =   27
      Top             =   4440
      Width           =   2055
      Begin VB.CommandButton Blower_Cur_AD 
         Caption         =   "Blower_Cur_AD"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton Blower_FB_AD 
         Caption         =   "Blower_FB_AD"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton P12V_VOL_AD 
         Caption         =   "P12V_VOL_AD"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Trig_Low_AD0 
         Caption         =   "SIG_Trig_Low_AD0"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Trig_Low_AD1 
         Caption         =   "SIG_Trig_Low_AD1"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   1695
      End
   End
   Begin VB.CommandButton DC_Motor_5V_AD 
      Caption         =   "DC_Motor_5V_AD"
      Height          =   375
      Left            =   2880
      TabIndex        =   26
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Location_DC_motor_AD0 
      Caption         =   "Location_DC_motor_AD0"
      Height          =   375
      Left            =   2880
      TabIndex        =   25
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Location_Step_AD0 
      Caption         =   "Location_Step_AD0"
      Height          =   375
      Left            =   2880
      TabIndex        =   24
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Andata2"
      Height          =   4095
      Left            =   2760
      TabIndex        =   18
      Top             =   120
      Width           =   2055
      Begin VB.CommandButton Location_DC_motor_AD1 
         Caption         =   "Location_DC_motor_AD1"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton Location_DC_motor_AD2 
         Caption         =   "Location_DC_motor_AD2"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Location_Step_AD1 
         Caption         =   "Location_Step_AD1"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Location_Step_AD2 
         Caption         =   "Location_Step_AD2"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton Location_DC_motor_AD3 
         Caption         =   "Location_DC_motor_AD3"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   1695
      End
   End
   Begin VB.CommandButton REQ_Gather_AD0 
      Caption         =   "REQ_Gather_AD0"
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton SIG_Sensor_AD9 
      Caption         =   "SIG_Sensor_AD9"
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton SIG_Sensor_AD8 
      Caption         =   "SIG_Sensor_AD8"
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton SIG_Trig_Hig_AD8 
      Caption         =   "SIG_Trig_Hig_AD8"
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Andata1"
      Height          =   4215
      Left            =   240
      TabIndex        =   9
      Top             =   4440
      Width           =   2055
      Begin VB.CommandButton SIG_Trig_Hig_AD6 
         Caption         =   "SIG_Trig_Hig_AD6"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Trig_Hig_AD7 
         Caption         =   "SIG_Trig_Hig_AD7"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Trig_Hig_AD5 
         Caption         =   "SIG_Trig_Hig_AD5"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton SIG_Trig_Hig_AD4 
         Caption         =   "SIG_Trig_Hig_AD4"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   1695
      End
   End
   Begin VB.CommandButton REQ_GATHER_AD5 
      Caption         =   "REQ_GATHER_AD5"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton REQ_GATHER_AD6 
      Caption         =   "REQ_GATHER_AD6"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton REQ_GATHER_AD7 
      Caption         =   "REQ_GATHER_AD7"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton REQ_GATHER_AD8 
      Caption         =   "REQ_GATHER_AD8"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Andata0"
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.CommandButton REQ_GATHER_AD2 
         Caption         =   "REQ_GATHER_AD2"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton REQ_GATHER_AD4 
         Caption         =   "REQ_GATHER_AD4"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton REQ_GATHER_AD1 
         Caption         =   "REQ_GATHER_AD1"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton REQ_GATHER_AD3 
         Caption         =   "REQ_GATHER_AD3"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'form2 按键每个对应一帧指令，指令有固定的帧头帧尾，指令须遵照规定的帧格式才能被下位机正确解析
Option Explicit
Dim t_frame(31) As Byte '发送的命令的帧
Dim t_head(1) As Byte   '帧头
Dim t_tail(1) As Byte   '帧尾

Dim checksum As Integer
Dim checksum2_4 As Integer

Private Sub Blower_Cur_AD_Click()
t_frame(5) = &H6
t_frame(6) = &H5
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub Blower_FB_AD_Click()
t_frame(5) = &H6
t_frame(6) = &H1
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub DC_Motor_5V_AD_Click()
t_frame(5) = &H1
t_frame(6) = &H4
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub Form_Load()
t_head(0) = &HFF    'PC上位机发送的固定帧头，两字节
t_head(1) = &H55
t_tail(0) = &HFF    'PC上位机发送的固定帧尾，两字节
t_tail(1) = &HAA

'AD类命令发送的部分固定格式
t_frame(0) = t_head(0)
t_frame(1) = t_head(1)
t_frame(2) = &H7
t_frame(3) = &HF0
t_frame(4) = &H2
't_frame(5)和t_frame(6)因采集的AD口不同而不同
't_frame(5) = &H3
't_frame(6) = &H0
checksum2_4 = t_frame(2) + t_frame(3) + t_frame(4)
'========下面注释的表示不能固化的部分，需要根据采集的AD做变动=============
'checksum = t_frame(2) + t_frame(3) + t_frame(4) + t_frame(5) + t_frame(6)
't_frame(7) = checksum \ 256
't_frame(8) = checksum Mod 256
t_frame(9) = t_tail(0)
t_frame(10) = t_tail(1)
End Sub

Private Sub KL15_Cur_AD_Click()
t_frame(5) = &H0
t_frame(6) = &H5
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub KL15_Vol_AD_Click()
t_frame(5) = &H0
t_frame(6) = &H0
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub KL30_Cur_AD_Click()
t_frame(5) = &H0
t_frame(6) = &H1
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub KL30_Vol_AD_Click()
t_frame(5) = &H3
t_frame(6) = &H3
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub Location_DC_motor_AD0_Click()
t_frame(5) = &H1
t_frame(6) = &H6
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub Location_DC_motor_AD1_Click()
t_frame(5) = &H1
t_frame(6) = &H5
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub Location_DC_motor_AD2_Click()
t_frame(5) = &H1
t_frame(6) = &H1
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub Location_DC_motor_AD3_Click()
t_frame(5) = &H1
t_frame(6) = &H3
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub Location_DC_motor_AD4_Click()
t_frame(5) = &H0
t_frame(6) = &H6
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub Location_Step_AD0_Click()
t_frame(5) = &H1
t_frame(6) = &H7
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub Location_Step_AD1_Click()
t_frame(5) = &H1
t_frame(6) = &H2
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub Location_Step_AD2_Click()
t_frame(5) = &H1
t_frame(6) = &H0
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub Location_Step_AD3_Click()
t_frame(5) = &H0
t_frame(6) = &H4
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub Location_Step_AD4_Click()
t_frame(5) = &H0
t_frame(6) = &H7
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub Motor_Step_Cur_AD_Click()
t_frame(5) = &H6
t_frame(6) = &H4
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub P12V_VOL_AD_Click()
t_frame(5) = &H6
t_frame(6) = &H2
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub POWER_Cur_AD_Click()
t_frame(5) = &H0
t_frame(6) = &H2
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub Power_Vol_AD_Click()
t_frame(5) = &H3
t_frame(6) = &H0
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub REQ_Gather_AD0_Click()
t_frame(5) = &H4
t_frame(6) = &H4
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub REQ_GATHER_AD1_Click()
t_frame(5) = &H5
t_frame(6) = &H2
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub REQ_GATHER_AD2_Click()
t_frame(5) = &H5
t_frame(6) = &H3
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub REQ_GATHER_AD3_Click()
t_frame(5) = &H5
t_frame(6) = &H1
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub REQ_GATHER_AD4_Click()
t_frame(5) = &H5
t_frame(6) = &H0
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub REQ_GATHER_AD5_Click()
t_frame(5) = &H5
t_frame(6) = &H4
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub REQ_GATHER_AD6_Click()
t_frame(5) = &H5
t_frame(6) = &H6
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub REQ_GATHER_AD7_Click()
t_frame(5) = &H5
t_frame(6) = &H7
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub REQ_GATHER_AD8_Click()
t_frame(5) = &H5
t_frame(6) = &H5
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub Reserved_AD0003_Click()
t_frame(5) = &H0
t_frame(6) = &H3
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub Reserved_AD0700_Click()
t_frame(5) = &H7
t_frame(6) = &H0
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub Reserved_AD0701_Click()
t_frame(5) = &H7
t_frame(6) = &H1
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub Reserved_AD0703_Click()
t_frame(5) = &H7
t_frame(6) = &H3
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Sensor_AD0_Click()
t_frame(5) = &H7
t_frame(6) = &H4
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Sensor_AD1_Click()
t_frame(5) = &H2
t_frame(6) = &H0
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Sensor_AD2_Click()
t_frame(5) = &H2
t_frame(6) = &H1
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Sensor_AD3_Click()
t_frame(5) = &H2
t_frame(6) = &H4
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Sensor_AD4_Click()
t_frame(5) = &H7
t_frame(6) = &H5
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Sensor_AD5_Click()
t_frame(5) = &H7
t_frame(6) = &H7
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Sensor_AD6_Click()
t_frame(5) = &H2
t_frame(6) = &H6
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Sensor_AD7_Click()
t_frame(5) = &H2
t_frame(6) = &H5
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Sensor_AD8_Click()
t_frame(5) = &H4
t_frame(6) = &H7
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Sensor_AD9_Click()
t_frame(5) = &H4
t_frame(6) = &H6
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Hig_AD0_Click()
t_frame(5) = &H3
t_frame(6) = &H7
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Hig_AD1_Click()
t_frame(5) = &H3
t_frame(6) = &H2
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Hig_AD2_Click()
t_frame(5) = &H3
t_frame(6) = &H1
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Hig_AD3_Click()
t_frame(5) = &H3
t_frame(6) = &H5
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Hig_AD4_Click()
t_frame(5) = &H4
t_frame(6) = &H3
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Hig_AD5_Click()
t_frame(5) = &H4
t_frame(6) = &H0
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Hig_AD6_Click()
t_frame(5) = &H4
t_frame(6) = &H1
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Hig_AD7_Click()
t_frame(5) = &H4
t_frame(6) = &H2
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Hig_AD8_Click()
t_frame(5) = &H4
t_frame(6) = &H5
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Hig_Cur_AD_Click()
t_frame(5) = &H6
t_frame(6) = &H7
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Low_AD0_Click()
t_frame(5) = &H6
t_frame(6) = &H0
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Low_AD1_Click()
t_frame(5) = &H6
t_frame(6) = &H3
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Low_AD2_Click()
t_frame(5) = &H3
t_frame(6) = &H4
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Low_AD3_Click()
t_frame(5) = &H3
t_frame(6) = &H6
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Low_AD4_Click()
t_frame(5) = &H2
t_frame(6) = &H7
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Low_AD5_Click()
t_frame(5) = &H2
t_frame(6) = &H2
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Low_AD6_Click()
t_frame(5) = &H2
t_frame(6) = &H3
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Low_AD7_Click()
t_frame(5) = &H7
t_frame(6) = &H2
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Low_AD8_Click()
t_frame(5) = &H7
t_frame(6) = &H6
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub

Private Sub SIG_Trig_Low_Cur_AD_Click()
t_frame(5) = &H6
t_frame(6) = &H6
checksum = checksum2_4 + t_frame(5) + t_frame(6)
t_frame(7) = checksum \ 256
t_frame(8) = checksum Mod 256
Form1.MSComm1.Output = t_frame

End Sub
