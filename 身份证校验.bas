Attribute VB_Name = "身份证校验"
'校验身份证并输出正确的身份证（矫正最后一位身份号）
Function MCID(ByVal oldIdStr As String) As String

'身份证前17位，每个数字系数
Dim factorData
ReDim oldIdArr(18)
Dim count
Dim lastId
Dim strLen
Dim gongshi
'当前正确身份证
Dim correctId
'系数求和
count = 0
factorData = Array(7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2)
verifyData = Array(1, 0, "X", 9, 8, 7, 6, 5, 4, 3, 2)
'去前面17,15位数字
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
'只负责判断是否正确 1正确 0失败
Function MCIDB(ByVal oldIdStr As String) As Integer

Dim rStatus As Integer

If MCID(oldIdStr) = oldIdStr And Len(oldIdStr) > 0 Then
    MCIDB = 1
Else
    MCIDB = 0
End If
End Function



