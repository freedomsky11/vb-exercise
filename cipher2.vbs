cipher = InputBox("ÇëÊäÈëÃÜÎÄ¡£")
plaintext = ""
For i = 1 To Len(cipher)
    letter = Mid(cipher, i, 1)
    plaintext = plaintext & Chr(asc(letter) - 3)
Next
MsgBox plaintext