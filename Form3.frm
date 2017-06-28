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
