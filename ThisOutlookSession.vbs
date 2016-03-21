Private Sub Application_Quit()
  If TimerID <> 0 Then Call DeactivateTimer 'Turn off timer upon quitting **VERY IMPORTANT**
End Sub

Private Sub Application_Startup()
  
  Call ActivateTimer(5) 'Set timer to go off every 1 minute
End Sub
