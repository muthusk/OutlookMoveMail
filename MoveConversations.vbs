Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerfunc As Long) As Long
Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Public TimerID As Long 'Need a timer ID to eventually turn off the timer. If the timer ID <> 0 then the timer is running

Public Sub TriggerTimer(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idevent As Long, ByVal Systime As Long)
  MoveConversation
End Sub


Public Sub DeactivateTimer()
Dim lSuccess As Long
  lSuccess = KillTimer(0, TimerID)
  If lSuccess = 0 Then
    MsgBox "The timer failed to deactivate."
  Else
    TimerID = 0
  End If
End Sub

Public Sub ActivateTimer(ByVal nMinutes As Long)
  nMinutes = nMinutes * 1000 * 60 'The SetTimer call accepts milliseconds, so convert to minutes
  If TimerID <> 0 Then Call DeactivateTimer 'Check to see if timer is running before call to SetTimer
  TimerID = SetTimer(0, 0, nMinutes, AddressOf TriggerTimer)
  If TimerID = 0 Then
    MsgBox "The timer failed to activate."
  End If
End Sub

Sub MoveConversation()

    Dim objOutlook As Outlook.Application
    Dim objNamespace As Outlook.NameSpace
    Dim objSourceFolder As Outlook.MAPIFolder
    Dim objDestFolder As Outlook.MAPIFolder
    Dim objVariant As Variant
    Dim lngMovedItems As Long
    Dim intCount As Integer
    Dim intDateDiff As Integer
    Dim strDestFolder As String
    
    Set objOutlook = Application
    Set objNamespace = objOutlook.GetNamespace("MAPI")
	
    Set objSourceFolder = objNamespace.Folders("CHANGEHTHIS@YOURDOMAIN.COM").Folders("SOURCE FOLDER NAME")
    
 
    Set objDestFolder = objNamespace.Folders("CHANGEHTHIS@YOURDOMAIN.COM").Folders("DESTINATION FODLER NAME")
    
    For intCount = objSourceFolder.Items.Count To 1 Step -1
        Set objVariant = objSourceFolder.Items.Item(intCount)
        DoEvents
        
        If objVariant.Class = olMail Then
            
             intDateDiff = DateDiff("d", objVariant.SentOn, Now)
             
            'Adjust as needed.
            If intDateDiff > 7 Then

              objVariant.Move objDestFolder
              
              'count the # of items moved
               lngMovedItems = lngMovedItems + 1

            End If
        End If
    Next
    
Set objDestFolder = Nothing
End Sub
