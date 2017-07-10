VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "开关界面"
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10650
   LinkTopic       =   "Form3"
   ScaleHeight     =   6675
   ScaleWidth      =   10650
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton SIGTrigHigEN0 
      Caption         =   "SIGTrigHigEN0"
      Height          =   615
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton SIGTrigLowEN3 
      Caption         =   "SIGTrigLowEN3"
      Height          =   615
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton SIGTrigLowEN2 
      Caption         =   "SIGTrigLowEN2"
      Height          =   615
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton SIGTrigLowEN1 
      Caption         =   "SIGTrigLowEN1"
      Height          =   615
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton SIGTrigLowEN0 
      Caption         =   "SIGTrigLowEN0"
      Height          =   615
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton HIGDriveCurEn 
      Caption         =   "HIGDriveCurEn"
      Height          =   615
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton BlowerRelayCtrlSW 
      Caption         =   "BlowerRelayCtrlSW"
      Height          =   615
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      Caption         =   "U10"
      Height          =   6255
      Left            =   8040
      TabIndex        =   27
      Top             =   240
      Width           =   2295
      Begin VB.CommandButton LowDriveCurEn 
         Caption         =   "LowDriveCurEn"
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton SWSIGSensor2 
      Caption         =   "SWSIGSensor2"
      Height          =   615
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton SWSIGSensor7 
      Caption         =   "SWSIGSensor7"
      Height          =   615
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton SWSIGSensor6 
      Caption         =   "SWSIGSensor6"
      Height          =   615
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton SWSIGSensor3 
      Caption         =   "SWSIGSensor3"
      Height          =   615
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton DarkCurEn 
      Caption         =   "DarkCurEn"
      Height          =   615
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton SIGTrigHigEN2 
      Caption         =   "SIGTrigHigEN2"
      Height          =   615
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton SIGTrigHigEN3 
      Caption         =   "SIGTrigHigEN3"
      Height          =   615
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "U9"
      Height          =   6255
      Left            =   5400
      TabIndex        =   18
      Top             =   240
      Width           =   2295
      Begin VB.CommandButton SIGTrigHigEN1 
         Caption         =   "SIGTrigHigEN1"
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "U8"
      Height          =   6255
      Left            =   2760
      TabIndex        =   9
      Top             =   240
      Width           =   2295
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton CleanOverCur 
      Caption         =   "CleanOverCur"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command1"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command1"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command1"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton SWSIGSensor4 
      Caption         =   "SWSIGSensor4"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton SWSIGSensor5 
      Caption         =   "SWSIGSensor5"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton SWSIGSensor1 
      Caption         =   "SWSIGSensor1"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "U7"
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      Begin VB.CommandButton SWSIGSensor0 
         Caption         =   "SWSIGSensor0"
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'form3 开关控制界面
Option Explicit
Dim t_frame(31) As Byte

Dim checksum As Integer
Dim checksum2_4 As Integer
Dim status As Long      '开关状态，因为采用了4块74HC595芯片级联，所以可同时控制32个开关
                        '需要注意的是status无法反映第31位开关的状态，因为数据溢出
Dim statusH As Byte     '储存第31位开关的状态
Dim power As Byte


Private Sub BlowerRelayCtrlSW_Click()
If BlowerRelayCtrlSW.Caption = "BlowerRelayCtrlSW" Then
    BlowerRelayCtrlSW.Caption = "CLOSE BlowerRelayCtrlSW"
    BlowerRelayCtrlSW.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
    power = 30
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
     BlowerRelayCtrlSW.Caption = "BlowerRelayCtrlSW"
     BlowerRelayCtrlSW.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 30
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub

Private Sub CleanOverCur_Click()
If CleanOverCur.Caption = "CleanOverCur" Then
    CleanOverCur.Caption = "CLOSE CleanOverCur"
    CleanOverCur.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
     
    power = 0
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
     CleanOverCur.Caption = "CleanOverCur"
     CleanOverCur.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 0
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub



Private Sub DarkCurEn_Click()
If DarkCurEn.Caption = "DarkCurEn" Then
    DarkCurEn.Caption = "CLOSE DarkCurEn"
    DarkCurEn.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
    power = 20
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
     DarkCurEn.Caption = "DarkCurEn"
     DarkCurEn.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 20
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub

Private Sub Form_Load()
status = 0          '上电默认32个开关全部处于关闭状态
'开关模块命令发送的部分固定格式
t_frame(0) = t_head(0)
t_frame(1) = t_head(1)
t_frame(2) = &H9
t_frame(3) = &HF0
t_frame(4) = &H4
't_frame(5) = &H0   't_frame(5) - t_frame(8)因控制的开关不同而不同
't_frame(6) = &H0
't_frame(7) = &H0
't_frame(8) = &H0
checksum2_4 = t_frame(2) + t_frame(3) + t_frame(4)
'========下面注释的表示不能固化的部分，需要根据控制的开关做变动=============
'checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
't_frame(9) = checksum \ 256
't_frame(10) = checksum Mod 256
t_frame(11) = t_tail(0)
t_frame(12) = t_tail(1)
End Sub

Private Sub HIGDriveCurEn_Click()
If HIGDriveCurEn.Caption = "HIGDriveCurEn" Then
    HIGDriveCurEn.Caption = "CLOSE HIGDriveCurEn"
    HIGDriveCurEn.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
    power = 29
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
     HIGDriveCurEn.Caption = "HIGDriveCurEn"
     HIGDriveCurEn.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 29
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub

Private Sub LowDriveCurEn_Click()
If LowDriveCurEn.Caption = "LowDriveCurEn" Then
    LowDriveCurEn.Caption = "CLOSE LowDriveCurEn"
    LowDriveCurEn.BackColor = RGB(0, 255, 0)      '绿色表示开关打开

'    power = 31
'    status = status Or (1 * 2 ^ power)
'因为2^31已经超出了LONG类型可以表示的最大整数范围，上句代码会引起数据溢出，所以这里做特别处理
    statusH = &H80
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
     LowDriveCurEn.Caption = "LowDriveCurEn"
     LowDriveCurEn.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
'    power = 31
'    status = status And (Not 1 * 2 ^ power)
    statusH = &H0
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub

Private Sub SIGTrigHigEN0_Click()
If SIGTrigHigEN0.Caption = "SIGTrigHigEN0" Then
    SIGTrigHigEN0.Caption = "CLOSE SIGTrigHigEN0"
    SIGTrigHigEN0.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
    power = 24
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
     SIGTrigHigEN0.Caption = "SIGTrigHigEN0"
     SIGTrigHigEN0.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 24
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub

Private Sub SIGTrigHigEN1_Click()
If SIGTrigHigEN1.Caption = "SIGTrigHigEN1" Then
    SIGTrigHigEN1.Caption = "CLOSE SIGTrigHigEN1"
    SIGTrigHigEN1.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
    power = 23
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
     SIGTrigHigEN1.Caption = "SIGTrigHigEN1"
     SIGTrigHigEN1.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 23
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub

Private Sub SIGTrigHigEN2_Click()
If SIGTrigHigEN2.Caption = "SIGTrigHigEN2" Then
    SIGTrigHigEN2.Caption = "CLOSE SIGTrigHigEN2"
    SIGTrigHigEN2.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
    power = 21
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
     SIGTrigHigEN2.Caption = "SIGTrigHigEN3"
     SIGTrigHigEN2.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 21
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub

Private Sub SIGTrigHigEN3_Click()
If SIGTrigHigEN3.Caption = "SIGTrigHigEN3" Then
    SIGTrigHigEN3.Caption = "CLOSE SIGTrigHigEN3"
    SIGTrigHigEN3.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
    power = 22
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
     SIGTrigHigEN3.Caption = "SIGTrigHigEN3"
     SIGTrigHigEN3.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 22
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub

Private Sub SIGTrigLowEN0_Click()
If SIGTrigLowEN0.Caption = "SIGTrigLowEN0" Then
    SIGTrigLowEN0.Caption = "CLOSE SIGTrigLowEN0"
    SIGTrigLowEN0.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
    power = 28
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
     SIGTrigLowEN0.Caption = "SIGTrigLowEN0"
     SIGTrigLowEN0.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 28
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub

Private Sub SIGTrigLowEN1_Click()
If SIGTrigLowEN1.Caption = "SIGTrigLowEN1" Then
    SIGTrigLowEN1.Caption = "CLOSE SIGTrigLowEN1"
    SIGTrigLowEN1.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
    power = 27
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
     SIGTrigLowEN1.Caption = "SIGTrigLowEN1"
     SIGTrigLowEN1.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 27
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub

Private Sub SIGTrigLowEN2_Click()
If SIGTrigLowEN2.Caption = "SIGTrigLowEN2" Then
    SIGTrigLowEN2.Caption = "CLOSE SIGTrigLowEN2"
    SIGTrigLowEN2.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
    power = 26
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
     SIGTrigLowEN2.Caption = "SIGTrigLowEN2"
     SIGTrigLowEN2.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 26
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub

Private Sub SIGTrigLowEN3_Click()
If SIGTrigLowEN3.Caption = "SIGTrigLowEN3" Then
    SIGTrigLowEN3.Caption = "CLOSE SIGTrigLowEN3"
    SIGTrigLowEN3.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
    power = 25
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
     SIGTrigLowEN3.Caption = "SIGTrigLowEN3"
     SIGTrigLowEN3.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 25
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub

Private Sub SWSIGSensor0_Click()
If SWSIGSensor0.Caption = "SWSIGSensor0" Then
    SWSIGSensor0.Caption = "CLOSE SWSIGSensor0"
    SWSIGSensor0.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
    power = 7
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
     SWSIGSensor0.Caption = "SWSIGSensor0"
     SWSIGSensor0.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 7
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub

Private Sub SWSIGSensor1_Click()
If SWSIGSensor1.Caption = "SWSIGSensor1" Then
    SWSIGSensor1.Caption = "CLOSE SWSIGSensor1"
    SWSIGSensor1.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
     
    power = 6
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
    
    SWSIGSensor1.Caption = "SWSIGSensor1"
    SWSIGSensor1.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 6
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub

Private Sub SWSIGSensor2_Click()
If SWSIGSensor2.Caption = "SWSIGSensor2" Then
    SWSIGSensor2.Caption = "CLOSE SWSIGSensor2"
    SWSIGSensor2.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
     
    power = 16
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
    
     SWSIGSensor2.Caption = "SWSIGSensor2"
     SWSIGSensor2.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 16
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub

Private Sub SWSIGSensor3_Click()
If SWSIGSensor3.Caption = "SWSIGSensor3" Then
    SWSIGSensor3.Caption = "CLOSE SWSIGSensor3"
    SWSIGSensor3.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
     
    power = 19
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
    
     SWSIGSensor3.Caption = "SWSIGSensor3"
     SWSIGSensor3.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 19
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub

Private Sub SWSIGSensor4_Click()
If SWSIGSensor4.Caption = "SWSIGSensor4" Then
    SWSIGSensor4.Caption = "CLOSE SWSIGSensor4"
    SWSIGSensor4.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
     
    power = 4
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
     SWSIGSensor4.Caption = "SWSIGSensor4"
     SWSIGSensor4.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 4
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub

Private Sub SWSIGSensor5_Click()
If SWSIGSensor5.Caption = "SWSIGSensor5" Then
    SWSIGSensor5.Caption = "CLOSE SWSIGSensor5"
    SWSIGSensor5.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
     
    power = 5
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
     SWSIGSensor5.Caption = "SWSIGSensor5"
     SWSIGSensor5.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 5
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub

Private Sub form_unload(cancel As Integer)
'退出时关闭所有开关
t_frame(5) = 0
t_frame(6) = 0
t_frame(7) = 0
t_frame(8) = 0
checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
t_frame(9) = checksum \ 256
t_frame(10) = checksum Mod 256

Form1.MSComm1.Output = t_frame
Form1.MSComm1.Output = vbCrLf
End Sub

Private Sub SWSIGSensor6_Click()
If SWSIGSensor6.Caption = "SWSIGSensor6" Then
    SWSIGSensor6.Caption = "CLOSE SWSIGSensor6"
    SWSIGSensor6.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
     
    power = 18
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
    
     SWSIGSensor6.Caption = "SWSIGSensor6"
     SWSIGSensor6.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 18
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub

Private Sub SWSIGSensor7_Click()
If SWSIGSensor7.Caption = "SWSIGSensor7" Then
    SWSIGSensor7.Caption = "CLOSE SWSIGSensor7"
    SWSIGSensor7.BackColor = RGB(0, 255, 0)      '绿色表示开关打开
    
     
    power = 17
    status = status Or (1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
Else
    
     SWSIGSensor7.Caption = "SWSIGSensor7"
     SWSIGSensor7.BackColor = RGB(222, 222, 222) '灰色表示开关关上
     
    power = 17
    status = status And (Not 1 * 2 ^ power)
    t_frame(5) = CByte(status \ (2 ^ 24)) Or statusH
    t_frame(6) = CByte((status \ (2 ^ 16)) And &HFF)
    t_frame(7) = CByte((status \ (2 ^ 8)) And &HFF)
    t_frame(8) = CByte(status And &HFF)
    checksum = checksum2_4 + t_frame(5) + t_frame(6) + t_frame(7) + t_frame(8)
    t_frame(9) = checksum \ 256
    t_frame(10) = checksum Mod 256
    
    Form1.MSComm1.Output = t_frame
End If
End Sub
