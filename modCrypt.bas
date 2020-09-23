Attribute VB_Name = "modCrypt"
#Const CASE_SENSITIVE_PASSWORD = False
'Encrypt text
Public Function EncryptText(strText As String, ByVal strPwd As String)
    Dim I As Integer, c As Integer
    Dim strBuff As String

#If Not CASE_SENSITIVE_PASSWORD Then

    'Convert password to upper case
    'if not case-sensitive
    strPwd = UCase$(strPwd)

#End If

    'Encrypt string
    If Len(strPwd) Then
        For I = 1 To Len(strText)
            c = Asc(Mid$(strText, I, 1))
            c = c + Asc(Mid$(strPwd, (I Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next I
    Else
        strBuff = strText
    End If
    EncryptText = strBuff
End Function

'Decrypt text encrypted with EncryptText
Public Function DecryptText(strText As String, ByVal strPwd As String)
    Dim I As Integer, c As Integer
    Dim strBuff As String

#If Not CASE_SENSITIVE_PASSWORD Then

    'Convert password to upper case
    'if not case-sensitive
    strPwd = UCase$(strPwd)

#End If

    'Decrypt string
    If Len(strPwd) Then
        For I = 1 To Len(strText)
            c = Asc(Mid$(strText, I, 1))
            c = c - Asc(Mid$(strPwd, (I Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next I
    Else
        strBuff = strText
    End If
    DecryptText = strBuff
End Function

