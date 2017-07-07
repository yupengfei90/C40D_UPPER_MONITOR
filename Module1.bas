Attribute VB_Name = "Module1"
Public t_frame(31) As Byte '发送的命令的帧
Public t_head(1) As Byte   '帧头
Public t_tail(1) As Byte   '帧尾
Public t_checksum As Long

