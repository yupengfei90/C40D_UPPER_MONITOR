Attribute VB_Name = "Module1"
'ȫ�ֱ��������д��ڣ�FORM�������ã�ҪС������֮����໥Ӱ��
'�粻��Ҫ���������̼佻�������ɶ����ڴ�

Public t_frame(31) As Byte '���͵������֡
Public t_head(1) As Byte   '֡ͷ
Public t_tail(1) As Byte   '֡β
Public t_checksum As Long

Public r_frame(31) As Byte '���յ������֡
Public r_head(1) As Byte   '֡ͷ
Public r_tail(1) As Byte   '֡β


'��ʱ50ms���ڼ���Խ��������л�
Public IsTimeElaplse As Byte

Public Sub Delay_50ms()
Form1.Timer2.Enabled = False
Form1.Timer2.Enabled = True
Do
    If IsTimeElaplse Then
        IsTimeElaplse = 0
        Exit Do
    End If
    DoEvents
Loop
End Sub

