Attribute VB_Name = "���֤У��"
'У�����֤�������ȷ�����֤���������һλ��ݺţ�
Function MCID(ByVal oldIdStr As String) As String

'���֤ǰ17λ��ÿ������ϵ��
Dim factorData
ReDim oldIdArr(18)
Dim count
Dim lastId
Dim strLen
Dim gongshi
'��ǰ��ȷ���֤
Dim correctId
'ϵ�����
count = 0
factorData = Array(7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2)
verifyData = Array(1, 0, "X", 9, 8, 7, 6, 5, 4, 3, 2)
'ȥǰ��17,15λ����
strLen = Len(oldIdStr) - 2
For i = 0 To strLen
    oldIdArr(i) = Mid(oldIdStr, i + 1, 1)
    count = count + oldIdArr(i) * factorData(i)
    correctId = correctId & oldIdArr(i)
Next

lastId = verifyData(count Mod 11)
correctId = correctId & lastId

MCID = correctId

End Function
'ֻ�����ж��Ƿ���ȷ 1��ȷ 0ʧ��
Function MCIDB(ByVal oldIdStr As String) As Integer

Dim rStatus As Integer

If MCID(oldIdStr) = oldIdStr And Len(oldIdStr) > 0 Then
    MCIDB = 1
Else
    MCIDB = 0
End If
End Function



