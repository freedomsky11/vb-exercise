k = InputBox("��������Կ��")
key = CInt(k)
text1 = InputBox("���������Ļ����ġ�")
text2 = ""
For i = 1 To Len(text1)
    letter = Mid(text1, i, 1)
    text2 = text2 & Chr(asc(letter) Xor key)
Next
MsgBox text2