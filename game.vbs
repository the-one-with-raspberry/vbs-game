Option Explicit

dim response
response = MsgBox("Press:" & vbCrLf & "Yes to continue" & vbCrLf & "No to go back" & vbCrLf & "Cancel to exit", 3, "Game")

Select Case response
Case vbYes: dim continuegame
    continuegame = MsgBox("Question: Is working in VBscript hard?", 4, "Game")
    If continuegame = vbYes Then
        Call MsgBox("WRONG!", 16, "Game")
    ElseIf continuegame = vbNo Then
        Call MsgBox("You passed the game! Yeah!", 64, "Game - Nice!")
    End If
Case vbNo: Call MsgBox("NULL EXCEPTION: NO BACK COMMAND: NEGATIVE NUMBER", 16, "Error")
End Select