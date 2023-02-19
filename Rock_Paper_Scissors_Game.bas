Attribute VB_Name = "Module2"

Sub gameRPS()

'User choice
'C7 is cell where user puts in his choice
userChoice = LCase(Range("c7").Value)

'Computer choice
compChoice = Application.WorksheetFunction.RandBetween(1, 3)

If compChoice = 1 Then
    compChoice = "rock"
ElseIf compChoice = 2 Then
    compChoice = "paper"
Else
    compChoice = "scissors"
End If

compChoice = LCase(compChoice)


'Comparing user and computer results
result = ""

If userChoice = compChoice Then
    result = "DRAW!"
ElseIf userChoice = "rock" And compChoice = "paper" Then
    result = "You lose - paper beats rock"
ElseIf userChoice = "rock" And compChoice = "scissors" Then
    result = "You won - rock beats scissors"
ElseIf userChoice = "paper" And compChoice = "rock" Then
    result = "You won - paper beats rock"
ElseIf userChoice = "paper" And compChoice = "scissors" Then
    result = "You lose - Scissors beats paper"
ElseIf userChoice = "scissors" And compChoice = "rock" Then
    result = "You lose - Rock beats scissors"
ElseIf userChoice = "scissors" And compChoice = "paper" Then
    result = "You won - Scissors beats paper"
End If

Debug.Print result

'C16 is cell where the result will be shown
Range("c16").Value = result



End Sub
