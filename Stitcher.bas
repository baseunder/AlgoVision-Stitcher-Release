'#Reference #System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL
'#Language "WWB.NET"
Dim stream As System.IO.stream
Dim webClient As New System.Net.WebClient

Sub CheckIfRunning()
    If System.Diagnostics.Process.GetProcessesByName("AlgoVision-Stitcher").Length  > 0 Or System.Diagnostics.Process.GetProcessesByName("Python").Length  > 0 Then
    	Debug.Print "isrunning"
    Else
      System.Diagnostics.Process.Start(MacroDir+"/AlgoVision-Stitcher.exe")
      Wait 3
    End If
End Sub
Sub Main
  CheckIfRunning()
  SC.A_SetLaserState(True)
	Begin Dialog UserDialog 400,203," ",.diafunc ' %GRID:10,7,1,1
		PushButton 20,21,150,77,"Mark",.PushButton1
		PushButton 210,21,150,77,"Setup Stitching",.PushButton3
		PushButton 30,133,160,56,"Stitch",.PushButton2
		PushButton 210,126,150,63,"Retrain Stitching",.PushButton4
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
End Sub

Sub stitch(ini As String)
  LC.SetBackgroundBitmapRef(New System.Drawing.Bitmap(10,10))
  Debug.Clear
  webClient.DownloadString("http://127.0.0.1:5001/"+ini)
  Dim bidir As Boolean = False
  Dim yt As Integer
  Dim stp As Integer = 8000
  Dim stb As Integer = CInt(stp*3.5)
  SC.S_Pos(-stb,-stb)
  Wait 0.01
  Dim count As Integer = 0
  Dim stpwtch As New System.Diagnostics.Stopwatch()
  LC.CallBackWhileRefresh(False)
  For x As Integer = -stb To stb Step stp
    For y As Integer = -stb To stb Step stp
      If bidir Then
        yt = -y
      Else
        yt = y
      End If
      SC.S_Pos(x,yt)
      SC.S_Pos(x,yt)
      webClient.DownloadString("http://127.0.0.1:5001/saveimage?index="+CStr(count))
      count = count + 1
    Next
    bidir = Not bidir
  Next
  LC.CallBackWhileRefresh(True)
End Sub

Rem Siehe DialogFunc Hilfethema für weitere Informationen.
Private Function diafunc(DlgItem As String, Action As Integer, SuppValue As Long) As Boolean
	Select Case Action
	Case 1 ' Dialogbox-Initialisierung
	Case 2 ' Wert verändert oder Schaltfläche gedrückt
		Rem diafunc = True ' Verhindert das Schließen des Dialogs beim Drücken der Schaltfläche
    DlgEnable("PushButton1",False)
    DlgEnable("PushButton2",False)
    DlgEnable("PushButton3",False)
    DlgEnable("PushButton4",False)
    If DlgItem.Equals("PushButton1") Then
      LC.Mark("",True)
    ElseIf DlgItem.Equals("PushButton2") Then
      stitch("initOpti")
      Wait 0.2
      stream = webClient.OpenRead("http://127.0.0.1:5001/stitch?part=live")
      LC.SetBackgroundBitmapRef(New System.Drawing.Bitmap(stream))
    ElseIf DlgItem.Equals("PushButton3") Then
      MsgBox("Load StitchCal.las and mark it on a tonerpaper, then confirm with OK to continue")
      stitch("init")
      Wait 0.2
      Dim uid =  CStr((Now - New System.DateTime(1970, 1, 1)).TotalMilliseconds())
      webClient.OpenRead("http://127.0.0.1:5001/save?part="+uid)
      webClient.OpenRead("http://127.0.0.1:5001/train?part="+uid)
      stream = webClient.OpenRead("http://127.0.0.1:5001/stitch?part=live")
      LC.SetBackgroundBitmapRef(New System.Drawing.Bitmap(stream))
    ElseIf DlgItem.Equals("PushButton4") Then
      webClient.OpenRead("http://127.0.0.1:5001/train?part=calibration")
      stream = webClient.OpenRead("http://127.0.0.1:5001/stitch?part=calibration")
      LC.SetBackgroundBitmapRef(New System.Drawing.Bitmap(stream))
    End If
    DlgEnable("PushButton1",True)
    DlgEnable("PushButton2",True)
    DlgEnable("PushButton3",True)
    DlgEnable("PushButton4",True)
    Return True
	Case 3 ' Text eines Textfelds oder eines Kombinationsfelds verändert
	Case 4 ' Fokus verändert
	Case 5 ' Leerlauf
		Wait .1 : diafunc = True ' Leerlaufaktionen weiterhin erhalten

	Case 6 ' Funktionstaste
	End Select
End Function
