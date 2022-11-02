Attribute VB_Name = "modProcessASOpen"

'@Folder("FixLinks2UDF")

Option Explicit

'==============================================================================
'modified version of 'modProcessWBOpen' from Jan Karel Pieterse from
'<https://www.jkp-ads.com/Articles/FixLinks2UDF.asp>
'Purpose (here): Delay execution of the 'ProcessAfterASOpen' macros until
'                there is an 'ActiveSheet'
'==============================================================================

'Counter to check how many times we've looped
Private mlTimesLooped As Long

Private Property Get TimesLooped() As Long
    TimesLooped = mlTimesLooped
End Property

Public Property Let TimesLooped(ByVal lTimesLooped As Long)
    mlTimesLooped = lTimesLooped
End Property

Public Sub CheckIfActiveSheetIsAvailable()
    If ActiveSheet Is Nothing Then
        RedoCheckIfActiveSheetIsAvailable
    Else
        'reset 'TimesLooped'
        TimesLooped = 0
        If Application.CalculationState = xlCalculating Then
            RedoCheckIfActiveSheetIsAvailable
        Else
            modProcessAfterASOpen.ProcessAfterASOpen
        End If
    End If
End Sub

Private Sub RedoCheckIfActiveSheetIsAvailable()
    'Increment the loop counter
    TimesLooped = TimesLooped + 1
    'May be needed if Excel is opened from a browser
    Application.Visible = True
    If TimesLooped < 20 Then
        '(to avoid a runtime error 1004 if a file is opened in "protected view",
        ' e.g. if a file from the internet is opened the first time)
        On Error Resume Next
        'We've not yet done this 20 times, schedule another in 1 sec
        Application.OnTime Now + TimeValue("00:00:01"), "CheckIfActiveSheetIsAvailable"
        On Error GoTo 0
    Else
        'We've done this 20 times, do not schedule another
        'and reset the counter
        TimesLooped = 0
    End If
End Sub
