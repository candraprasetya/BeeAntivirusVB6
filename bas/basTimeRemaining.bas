Attribute VB_Name = "basTimeRemaining"
'**********************
'
'   Time Remaining Calculator
'
'    by Kendall Cain
'    gabberhed@usa.net
'
'   This calculates the average time taken
'   for a progressbar to get to its current
'   percentage and calculates how much time
'   it will take for the progressbar to
'   reach 100%
'
'   All you do is:
'   1.) Declare a variable in the (General)
'       area of your form as a string
'   2.) When the progressbar starts call 'BeginProgress'
'   3.) When ever you make the progressbars value increase
'       get the percentage with the 'getPercentage' function
'       then call the 'TimeRemaining' function
'   4.) Voila! You will have the approxamate remaining time
'       for your progressbar.
'
'  *note: If you get your stopwatch out and see that it isnt
'         going second by second, its just like a download
'         time calculator, the closer you get to 100% the
'         more accurate it gets, also it depends on the
'         programmers competentcy to increment their progressbar
'         equally from 0 to 100. Meaning you dont increase it by
'         1 percent for some small process and then 10 percent
'         for an equally small process.   =)
'         Enjoy!
'**********************
Dim sStart As String

Function BeginProgress()
'Call this when you begin th progressbar
sStart = Now
End Function

Function TimeRemaining(CurrentPercentage As String) As String

Dim maxTime As String
Dim sRemaining As String
Dim Mins As String

'On Error GoTo Handler

'determine how much time(in seconds) has passed since the progress bar has started
maxTime = DateDiff("s", sStart, Now)

'make sure percentage is above 0%
If CurrentPercentage > 0 Then
    'calculate how many seconds until progressbar is finished
    sRemaining = Val(maxTime / CurrentPercentage) * Val(100 - CurrentPercentage)
    'convert seconds into Minutes:Seconds format
    Mins = Format(Fix(sRemaining / 60), "00")
    'set return variable to have Minutes:Seconds left and also Hours
    TimeRemaining = Format(Fix(Mins / 60), "00") & ":" & Format(Mins Mod 60, "00") & ":" & Format(sRemaining Mod 60, "00")
End If

Exit Function

Handler:
TimeRemaining = "Error"
MsgBox "Error #" & Err.Number & vbCrLf & Err.Description, vbOKOnly, "Error"

End Function

Function getPercentage(ProgressBarCurrentValue As String, ProgressBarMaxValue As String) As String
'calculate Percentage completed of progressbar
getPercentage = Format(Val(Val(ProgressBarCurrentValue / ProgressBarMaxValue) * 100), "0")
End Function

