Attribute VB_Name = "Module1"
'全局变量，所有窗口（FORM）均可用，要小心它们之间的相互影响
'如不需要在整个工程间交互，不可定义在此

Public t_frame(31) As Byte '发送的命令的帧
Public t_head(1) As Byte   '帧头
Public t_tail(1) As Byte   '帧尾
Public t_checksum As Long

Public r_frame(31) As Byte '接收的命令的帧
Public r_head(1) As Byte   '帧头
Public r_tail(1) As Byte   '帧尾

