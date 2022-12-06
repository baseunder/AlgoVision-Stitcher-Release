'#Reference #System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL
'#Language "WWB.NET"

Dim pixelToUm As Double = 71.6

Dim stream As System.IO.Stream
Dim webClient As New System.Net.WebClient
Dim stp As Integer = 8000
Dim offX As Integer = 0
Dim offY As Integer = 0
Dim srcWidth As Integer = 1600
Dim srcHeight As Integer = 1600
Dim tBtm As System.Drawing.Bitmap
Dim smallView As Boolean = False
Dim parts As New System.Collections.Generic.List(Of part)
Dim curPatternName As String
Dim isInMark As Boolean = False

Dim MATERIAL As String
Dim FK As String

Structure part
Dim x As Integer
Dim y As Integer
Dim a As Double
End Structure

Sub CheckIfRunning()
	If System.Diagnostics.Process.GetProcessesByName("AlgoVision-Stitcher").Length  > 0 Then
			Debug.Print "stitcher is running"
	Else
		If Not System.Diagnostics.Process.GetProcessesByName("Python").Length  > 0 Then
			System.Diagnostics.Process.Start(MacroDir+"/AlgoVision-Stitcher.exe")
			Wait 3
		End If
	End If
	If System.Diagnostics.Process.GetProcessesByName("AlgoVision-Matcher").Length  > 0 Then
			Debug.Print "matcher is running"
	Else
		System.Diagnostics.Process.Start(MacroDir+"/AlgoVision-Matcher.exe")
	End If
End Sub

Sub loadFieldSize()
	Dim fslines() As String = System.IO.File.ReadAllLines(MacroDir+"\field.size")
	offX = CInt(fslines(0))
	offY = CInt(fslines(1))
	stp = CInt(fslines(2))
	Debug.Print "field.size file found and loaded"
End Sub

Dim buttonList As New System.Collections.Generic.List(Of String)
Sub Main
	Debug.Clear
	CheckIfRunning()
	SC.A_SetLaserState(True)
	If System.IO.File.Exists(MacroDir+"\field.size") Then
		loadFieldSize()
	Else
		MsgBox("No field.size file present, please generate one with the MiddleCam.bas Skript",VbMsgBoxStyle.vbOkOnly, "Use standard values")
	End If
	buttonList.Add("PushButton1")
	buttonList.Add("PushButton2")
	buttonList.Add("PushButton3")
	buttonList.Add("PushButton4")
	buttonList.Add("PushButton5")
	buttonList.Add("PushButton6")
	buttonList.Add("PushButton7")
	buttonList.Add("PushButton8")
	buttonList.Add("PushButton9")
	buttonList.Add("PushButton10")
	buttonList.Add("PushButtonPreview")
	If LC.GetPW_Level()<3 Then
	Begin Dialog UserDialog 990,238," ",.diafunc ' %GRID:10,7,1,1
		GroupBox 440,21,220,210,"Experimental",.GroupBox1
		PushButton 470,42,170,35,"Reload camConf.txt",.PushButton5
		PushButton 470,77,170,35,"Retrain Stitching",.PushButton4
		GroupBox 20,28,380,196,"Ready to work",.GroupBox2
		PushButton 240,49,140,35,"Match Pattern",.PushButton7
		PushButton 30,133,210,77,"Stitch Image + Match Pattern",.PushButton9
		PushButton 240,84,140,35,"Preview Part",.PushButtonPreview
		PushButton 30,49,180,35,"Stitch",.PushButton2
		PushButton 30,91,180,35,"Load",.PushButton10
		PushButton 240,133,140,77,"Mark",.PushButton1
		GroupBox 690,21,240,98,"Setup",.GroupBox3
		GroupBox 700,126,240,105,"Vision",.GroupBox4
		PushButton 710,77,170,35,"Get Mid Point",.PushButton8
		PushButton 710,42,120,35,"Setup Stitching",.PushButton3
		TextBox 840,147,80,14,.TextBoxPiece
		PushButton 460,168,180,56,"Teach Pattern",.PushButton6
		TextBox 840,168,80,14,.TextBoxOverlap
		TextBox 840,189,80,14,.TextBoxSimilarity
		Text 720,147,100,14,"max. Piece",.Text1
		Text 720,168,100,14,"max. Overlap",.Text2
		Text 720,189,100,14,"min. Similarity",.Text3
		CheckBox 720,210,100,14,"calc Angle",.angleCheckBox
	End Dialog
		Dim dlg As UserDialog
		Dialog dlg
	Else
		Begin Dialog UserDialog 420,250," ",.diafunc ' %GRID:10,7,1,1
		GroupBox 440,21,220,210,"Experimental",.GroupBox1
		PushButton 470,42,170,35,"Reload camConf.txt",.PushButton5
		PushButton 470,77,170,35,"Retrain Stitching",.PushButton4
		PushButton 470,112,170,35,"Get Mid Point",.PushButton8
		GroupBox 20,28,380,196,"Ready to work",.GroupBox2
		PushButton 240,49,140,35,"Match Pattern",.PushButton7
		PushButton 30,133,210,77,"Stitch Image + Match Pattern",.PushButton9
		PushButton 240,84,140,35,"Preview Part",.PushButtonPreview
		PushButton 30,49,140,35,"Stitch",.PushButton2
		PushButton 30,91,180,35,"Load",.PushButton10
		PushButton 240,133,140,77,"Mark",.PushButton1
		GroupBox 690,21,240,98,"Setup",.GroupBox3
		GroupBox 700,126,240,105,"Vision",.GroupBox4
		PushButton 710,42,120,35,"Setup Stitching",.PushButton3
		TextBox 840,147,80,14,.TextBoxPiece
		PushButton 460,168,180,56,"Teach Pattern",.PushButton6
		TextBox 840,168,80,14,.TextBoxOverlap
		TextBox 840,189,80,14,.TextBoxSimilarity
		Text 720,147,100,14,"max. Piece",.Text1
		Text 720,168,100,14,"max. Overlap",.Text2
		Text 720,189,100,14,"min. Similarity",.Text3
		CheckBox 720,210,100,14,"calc Angle",.angleCheckBox
		End Dialog
		Dim dlg2 As UserDialog
			Dialog dlg2
	End If
End Sub

Sub calcField()
	Dim FieldSize As Integer
	Dim xCorr As Integer
	Dim yCorr As Integer
	Debug.Clear
	LC.LoadFile(MacroDir+"\MiddlePoint.LAS")
	Debug.Print "Mark a small and filled circle on the 0/0 coordinates onto a tonerpaper"
	If MsgBox("Should i mark the currently loaded MiddlePoint.LAS file?",VbMsgBoxStyle.vbYesNo,"Mark MiddlePoint") = vbYes Then
		LC.Mark("",True)
	End If
	SC.A_SetPilotState(False)
	LC.SetFocusFinderState(False)
	webClient.DownloadString("http://127.0.0.1:5001/init")
	Dim stepsize As Integer = 1500
	Dim counter As Integer = 1
  SC.P_SetOffsetPilot(0,0)
  SC.P_SetSizeFactorPilot(1,1)
	SC.S_Pos(0,0)
	Wait 0.1
	Dim re1() As String = Split(webClient.DownloadString("http://127.0.0.1:5001/holdImage"),";")
	stream = webClient.OpenRead("http://127.0.0.1:5001/fullImage")
	LC.SetBackgroundBitmapRef(New System.Drawing.Bitmap(stream))
  SC.P_SetOffsetPilot(0,0)
  SC.P_SetSizeFactorPilot(1,1)
	SC.S_Pos(stepsize,0)
	Wait 0.1
	Dim re2() As String = Split(webClient.DownloadString("http://127.0.0.1:5001/holdImage"),";")
  SC.P_SetOffsetPilot(0,0)
  SC.P_SetSizeFactorPilot(1,1)
	SC.S_Pos(0,stepsize)
	Wait 0.1
	Dim re3() As String = Split(webClient.DownloadString("http://127.0.0.1:5001/holdImage"),";")
	Dim xStp As Integer = re2(0)-re1(0)
	Dim yStp As Integer =  re3(1)-re1(1)
	Dim xDist As Integer = (CInt(re1(0))-CInt(CInt(re1(2))/2))
	Dim yDist As Integer = (CInt(re1(1))-CInt(CInt(re1(3))/2))
	xCorr = CInt((stepsize/-xStp)*xDist)
	yCorr = CInt((stepsize/-yStp)*yDist)
	Dim xFieldSize As Integer = Abs(CInt(((CInt(re1(2))*0.75)/xStp)*stepsize))
	Dim yFieldSize As Integer = Abs(CInt(((CInt(re1(3))*0.75)/yStp)*stepsize))
	FieldSize = xFieldSize
	If yFieldSize < FieldSize Then FieldSize = yFieldSize
	FieldSize = FieldSize - (FieldSize Mod 100)
	Debug.Print "Correction X/Y in um:"
	Debug.Print xCorr
	Debug.Print yCorr
	Debug.Print "Fieldsize:"
	Debug.Print FieldSize
  SC.P_SetOffsetPilot(0,0)
  SC.P_SetSizeFactorPilot(1,1)
	SC.S_Pos(xCorr,yCorr)
	Wait 0.1
	stream = webClient.OpenRead("http://127.0.0.1:5001/fullImage")
	LC.SetBackgroundBitmapRef(New System.Drawing.Bitmap(stream))
	System.IO.File.WriteAllText(MacroDir+"\field.size",CStr(xCorr)+vbNewLine+CStr(yCorr)+vbNewLine+CStr(FieldSize))
	loadFieldSize()
	LC.SaveFile("t.t")
	LC.LoadFile("")
	LC.TreeViewAddEOM("Header","MATRIX")
	LC.TreeViewAddEOM("Ellipse","CIRCLE")
	LC.SetNumericValue("MATRIX",12,CInt(-(4*FieldSize)))
	LC.SetNumericValue("MATRIX",13,CInt(-(4*FieldSize)))
	LC.SetNumericValue("MATRIX",19,FieldSize)
	LC.SetNumericValue("MATRIX",21,FieldSize)
	LC.SetNumericValue("MATRIX",20,9)
	LC.SetNumericValue("MATRIX",22,9)
	LC.SetBooleanValue("MATRIX",7,True)
	LC.SetBooleanValue("MATRIX",9,True)
	LC.SetBooleanValue("MATRIX",10,True)
	LC.SetNumericValue("MATRIX",1,30)
	LC.SetNumericValue("MATRIX",2,500000)
	LC.SetNumericValue("MATRIX",3,30000)
	LC.SetNumericValue("MATRIX",4,0.161)
	LC.SetNumericValue("MATRIX",29,81)
	LC.SetNumericValue("CIRCLE",4,250)
	LC.SetNumericValue("CIRCLE",5,250)
	LC.SetNumericValue("CIRCLE",7,720)
	LC.SaveFile(MacroDir+"\StitchCalFieldSize.LAS")
	MsgBox("Set your um/px values [Programm Settings, pixelToUm in Skript] to FIELDSIZE/SUBIMAGESIZE (8700/80)=108.75")
End Sub

Dim isInStitch As Boolean
Sub stitch(ini As String)
  isInStitch = True
  'LC.CallBackWhileRefresh(False)
	Debug.Print "start stitching"
	LC.SetBackgroundBitmapRef(New System.Drawing.Bitmap(10,10))
	webClient.DownloadString("http://127.0.0.1:5001/"+ini)
	USC_WS.SetOut1(True)
	USC_WS.SetOut2(True)
	SC.A_SetPilotState(False)
	webClient.OpenRead("http://127.0.0.1:5001/startGrab")
	Dim bidir As Boolean = False
	Dim yt As Integer
	Dim stb As Integer = CInt(stp*3.5)
  SC.P_SetOffsetPilot(0,0)
  SC.P_SetSizeFactorPilot(1,1)
	SC.S_Pos(-stb,-stb)
	Wait 0.01
	Dim count As Integer = 0
	Dim stpwtch As New System.Diagnostics.Stopwatch()
	Dim Lt As New System.Collections.Generic.List(Of Integer)()
	If smallView Then
		Lt.Add(18)
		Lt.Add(19)
		Lt.Add(20)
		Lt.Add(21)
		
		Lt.Add(26)
		Lt.Add(27)
		Lt.Add(28)
		Lt.Add(29)
		
		Lt.Add(34)
		Lt.Add(35)
		Lt.Add(36)
		Lt.Add(37)
		
		Lt.Add(42)
		Lt.Add(43)
		Lt.Add(44)
		Lt.Add(45)
	Else
		'For o As Integer = 16 To 47
		'	Lt.Add(o)
		'Next
		For o As Integer = 0 To 63
		  Lt.Add(o)
	  Next
	End If
	
  SC.P_SetOffsetPilot(0,0)
  SC.P_SetSizeFactorPilot(1,1)
	For x As Integer = -stb To stb Step stp
		For y As Integer = -stb To stb Step stp
			If bidir Then
				yt = -y
			Else
				yt = y
			End If
			If Lt.Contains(count) Then
				SC.S_Pos(x+offX,yt+offY)
				System.Threading.Thread.Sleep(35)
				If count.Equals(0) Then System.Threading.Thread.Sleep(100)
				webClient.DownloadString("http://127.0.0.1:5001/saveimage?index="+CStr(count)+"&comp=0")
			End If
			count = count + 1
		Next
		bidir = Not bidir
	Next
	'LC.CallBackWhileRefresh(True)
	webClient.OpenRead("http://127.0.0.1:5001/stopGrab")
	USC_WS.SetOut1(False)
	USC_WS.SetOut2(False)
	SC.A_SetPilotState(True)
	Debug.Print "stop stitching"
  SC.S_Pos(0,0)
  isInStitch = False
End Sub

Sub loadParts(inp As String, nbr As Integer)
	curPatternName = ""
	If nbr < 0 Then
		parts.Clear()
		Debug.Print inp
		Dim inpsp() As String = Split(inp,"#")
		For Each inps As String In inpsp
			If inps.Length = 0 Then
				Exit For
			End If
			Dim spl() As String = Split(inps,";")
			Dim p As part
			p.x = CInt((CInt(srcWidth/2)-CInt(spl(3)))*-pixelToUm)
			p.y = CInt((CInt(srcHeight/2)-CInt(spl(4)))*pixelToUm)
			p.a = CDbl(Replace(spl(2),".",","))
			parts.Add(p)
		Next
		If parts.Count() = 0 Then
			LC.SetCheck("VISION_POS",False)
			Return
		End If
		nbr = 0
		Debug.Print CStr(parts.Count)+" parts loaded"
		curPatternName = inpsp(inpsp.Length -1)
		Debug.Print curPatternName
	End If
	Dim zt As Integer
	LC.SetCheck("VISION_POS",True)
	If Not LC.GetCheck("VISION_POS") Then
		LC.SaveFile(MacroDir+"\t.t")
		LC.LoadFile(MacroDir+"\Layouts\generic.LAS")
	End If
  LC.CallBackWhileRefresh(False)
  LC.SetNumericValue("VISION_POS",20,parts.Count())
  LC.SetNumericValue("VISION_POS",29,parts.Count())
  LC.CallBackWhileRefresh(True)
	If Not isInMark Then
		LC.SetNumericValue("VISION_POS",12,parts(nbr).x)
		LC.SetNumericValue("VISION_POS",13,parts(nbr).y)
		LC.SetNumericValue("VISION_POS",23,parts(nbr).x)
		LC.SetNumericValue("VISION_POS",24,parts(nbr).y)
		LC.SetNumericValue("VISION_POS",25,parts(nbr).a)
		If Not System.IO.Directory.Exists(MacroDir+"\Layouts") Then
			System.IO.Directory.CreateDirectory(MacroDir+"\Layouts")
		End If
		If curPatternName.Length > 0 Then
			LC.SaveFile(MacroDir+"\Layouts\"+curPatternName+"_"+visionParams("iMaxPos")+"_"+visionParams("dScore") +"_myPart.LAS")
		End If
		LC.Refresh()
	End If
End Sub

Dim visionParams As New System.Collections.Generic.Dictionary(Of String, String)
Function updateParams() As String
	visionParams.Clear()
	visionParams.Add("iMaxPos",DlgText "TextBoxPiece")
	visionParams.Add("dMaxOverlap",DlgText "TextBoxOverlap")
	visionParams.Add("dScore",DlgText "TextBoxSimilarity")
	If DlgValue("angleCheckBox") Then
		visionParams.Add("dToleranceAngle", "180")
	Else
		visionParams.Add("dToleranceAngle", "0")
	End If
	Dim res As String
	For Each key As String In visionParams.Keys
		res += key+"="+visionParams(key)+"&"
	Next
'Debug.Print res
	Return res
End Function
Sub runStitch()
	LC.TimerStart()
	stitch("initOpti")
	stream = webClient.OpenRead("http://127.0.0.1:5001/stitch?part=live")
	LC.SetBackgroundBitmapRef(New System.Drawing.Bitmap(stream))
	LC.TimerStop()
End Sub
Sub runMatch()
	webClient.DownloadString("http://127.0.0.1:5001/updateVisionParams?"+updateParams())
	Try
	loadParts webClient.DownloadString("http://127.0.0.1:5001/update?part=live"),-1
	Catch
	MsgBox("no parts found")
	Finally
	End Try
	stream = webClient.OpenRead("http://127.0.0.1:5001/getMatchRes")
	tBtm = New System.Drawing.Bitmap(stream)
	srcWidth = tBtm.Width
	srcHeight = tBtm.Height
	LC.SetBackgroundBitmapRef(tBtm)
End Sub

Rem Siehe DialogFunc Hilfethema für weitere Informationen.
Private Function diafunc(DlgItem As String, Action As Integer, SuppValue As Long) As Boolean
	Select Case Action
	Case 1 ' Dialogbox-Initialisierung
		Try
		DlgText "TextBoxPiece", "1"
		DlgText "TextBoxOverlap", "0"
		DlgText "TextBoxSimilarity", "0.6"
		DlgValue "angleCheckBox", "1"
		Catch
		Finally
		End Try
		updateParams()
		curFileName = System.IO.Path.GetFileName(LC.GetFileName())
		parseFileName(System.IO.Path.GetFileName(LC.GetFileName()))
	Case 2 ' Wert verändert oder Schaltfläche gedrückt
		Rem diafunc = True ' Verhindert das Schließen des Dialogs beim Drücken der Schaltfläche
		For Each button As String In buttonList
			DlgEnable(button,False)
		Next
		
		If DlgItem.Equals("PushButton1") Then
      WORKSTATION.Door(False)
      wsWait()
      Wait 0.5
			SC.A_SetPilotState(True)
			LC.CallBackWhileRefresh(False)
			isInMark = True
      SC.A_SetShutterState(True)
      Wait 0.2
			LC.Mark("",False)
      SC.A_SetShutterState(False)
			isInMark = False
			LC.CallBackWhileRefresh(True)
			SC.A_SetPilotState(False)
      WORKSTATION.Door(True)
      ztMarked = -1
		ElseIf DlgItem.Equals("PushButton2") Then ' stitch
			runStitch()
		ElseIf DlgItem.Equals("PushButton3") Then
			MsgBox("Load StitchCalFieldSize.las and mark it on a tonerpaper, then confirm with OK to continue")
			stitch("init")
			Wait 0.2
			Dim uid =  CStr((Now - New System.DateTime(1970, 1, 1)).TotalMilliseconds())
			webClient.OpenRead("http://127.0.0.1:5001/save?part="+uid)
			webClient.OpenRead("http://127.0.0.1:5001/train?part="+uid)
			stream = webClient.OpenRead("http://127.0.0.1:5001/stitch?part="+uid)
			LC.SetBackgroundBitmapRef(New System.Drawing.Bitmap(stream))
		ElseIf DlgItem.Equals("PushButton4") Then ' retrain calib
			Dim calibName As String = InputBox("Enter the calib folder name","","calibration")
			webClient.OpenRead("http://127.0.0.1:5001/train?part="+calibName)
			stream = webClient.OpenRead("http://127.0.0.1:5001/stitch?part="+calibName)
			LC.SetBackgroundBitmapRef(New System.Drawing.Bitmap(stream))
		ElseIf DlgItem.Equals("PushButton5") Then ' reload config
			webClient.DownloadString("http://127.0.0.1:5001/loadConfig")
		ElseIf DlgItem.Equals("PushButton6") Then ' teach pattern
			webClient.DownloadString("http://127.0.0.1:5001/updateVisionParams?"+updateParams())
			loadParts webClient.DownloadString("http://127.0.0.1:5001/selectPattern?part=live"),-1
			stream = webClient.OpenRead("http://127.0.0.1:5001/getMatchRes")
			tBtm = New System.Drawing.Bitmap(stream)
			srcWidth = tBtm.Width
			srcHeight = tBtm.Height
			LC.SetBackgroundBitmapRef(tBtm)
		ElseIf DlgItem.Equals("PushButton7") Then ' match
			runMatch()
		ElseIf DlgItem.Equals("PushButton9") Then ' stitch and match
			diafunc("PushButton2",2,0)
			diafunc("PushButton7",2,0)
		ElseIf DlgItem.Equals("PushButton8") Then
			calcField()
		ElseIf DlgItem.Equals("PushButtonPreview") Then
			If parts.Count() > 0 Then
				previewCount = previewCount + 1
				previewCount = previewCount Mod parts.Count()
				loadParts "",previewCount
			Else
				MsgBox("no parts found")
			End If
		ElseIf DlgItem.Equals("PushButton10") Then
      MATERIAL = InputBox("Bitte Material abscannen","Material")
      FK = InputBox("Bitte FK abscannen","FK")
      Dim filelist() As String = System.IO.Directory.GetFiles(MacroDir+"\Layouts\")
      Dim foundFile As String
      For Each fln As String In filelist
        Debug.Print fln
        If System.IO.Path.GetFileName(fln).Contains(MATERIAL) Then
          foundFile = fln
        End If
      Next
      If foundFile.Length > 0 Then
        LC.LoadFile(foundFile)
      Else
        MsgBox("Layout not found")
        MATERIAL = ""
        FK = ""
      End If
		End If
		For Each buttone As String In buttonList
			DlgEnable(buttone,True)
		Next
		
		Return True
	Case 3 ' Text eines Textfelds oder eines Kombinationsfelds verändert
	Case 4 ' Fokus verändert
	Case 5 ' Leerlauf
		Wait .1 : diafunc = True ' Leerlaufaktionen weiterhin erhalten
		If reloadParams Then
			reloadParams = False
			DlgText "TextBoxPiece", visionParams("iMaxPos")
			DlgText "TextBoxSimilarity", visionParams("dScore")
		End If
	Case 6 ' Funktionstaste
	End Select
End Function
Dim reloadParams As Boolean = False
Dim previewCount As Integer = 0

Public Sub LC_ElementExit(ByRef sLabel As String) Handles LC.ElementExit
	Debug.Print "EXIT "+sLabel
End Sub
Dim curFileName As String


Public Sub LC_LasFileChanged(ByRef sReason As String, ByRef sFullName As String, ByRef sDirectoryName As String, ByRef sFileName As String) Handles LC.LasFileChanged
	If Not curFileName.Equals(sFileName) Then
		curFileName = sFileName
		parseFileName(sFileName)
	End If
End Sub

Sub parseFileName(sFileName As String)
	Dim splts() As String = Split(sFileName,"_")
	If splts.Length>1 Then
		If System.IO.File.Exists(MacroDir+"\templates\"+splts(0)+".bmp") Then
			webClient.DownloadString("http://127.0.0.1:5001/loadPattern?pattern="+splts(0))
		End If
	End If
	If splts.Length>2 Then
		visionParams("iMaxPos") = splts(1)
		visionParams("dScore") = splts(2)
		reloadParams = True
	End If
End Sub
Dim zt As Integer
Public Sub LC_ElementEntry(ByRef sLabel As String, ByRef lArrayX As Integer, ByRef lArrayY As Integer) Handles LC.ElementEntry
	Debug.Print "ENTRY "+sLabel +" "+ CStr(lArrayX)
	Try
	If isInMark Then 'And Not isInStitch Then
		If sLabel.Equals("VISION_CAM") Then ' And Not ztMarked.Equals(zt) Then
			Debug.Print("perform stitch and match")
			runStitch()
			runMatch()
			LC.SetNumericValue("VISION_POS",12,0)
			LC.SetNumericValue("VISION_POS",13,0)
			LC.SetNumericValue("VISION_POS",23,0)
			LC.SetNumericValue("VISION_POS",24,0)
			LC.SetNumericValue("VISION_POS",25,0)
		ElseIf sLabel.Equals("VISION_POS") Then
			zt = lArrayX
			Debug.Print sLabel+" "+ CStr(zt)
		End If
	End If
	Catch
	Finally
	End Try
End Sub
Dim ztMarked As Integer =-1
Public Sub LC_ElementBeforeLaser(ByRef sLabel As String, ByRef lArrayX As Integer, ByRef lArrayY As Integer) Handles LC.ElementBeforeLaser
	If zt < parts.Count() Then
		Debug.Print CStr(Now()) +" "+ CStr(zt) +" "+sLabel
		SC.A_SetPilotState(True)
		HG.Move(parts(zt).x, parts(zt).y)
		HG.Rotation(parts(zt).x, parts(zt).y,parts(zt).a)
    ztMarked = zt
	End If
End Sub

Public Sub LC_MarkingEnd(ByRef iMarkStatus As Integer) Handles LC.MarkingEnd
	isInMark = False
End Sub

Public Sub LC_OnError() Handles LC.OnError
	
End Sub

Public Sub LC_OnMessage(ByRef sDate As String, ByRef sTime As String, ByRef lID As Long, ByRef sMessage As String) Handles LC.OnMessage
	
End Sub

Public Sub LC_ShowMarkChanged(ByRef lIndex As Integer, ByRef lX1 As Integer, ByRef lY1 As Integer, ByRef lX2 As Integer, ByRef lY2 As Integer) Handles LC.ShowMarkChanged
	
End Sub

Function getSerial(bMark As Boolean) As String
  Dim sndir As String = MacroDir+"\Serials"
  If Not System.IO.Directory.Exists(sndir) Then
    System.IO.Directory.CreateDirectory(sndir)
  End If
  Dim fil As String = sndir+"\"+FK+".txt"
  Dim zeroLen As Integer
  Dim curSerial As Integer = -1
  If System.IO.File.Exists(fil) Then
    Debug.Print "serial file found"
    Dim lines() As String = System.IO.File.ReadAllLines(fil)
    zeroLen = CInt(lines(0))
    curSerial = CInt(lines(1))
  Else
    zeroLen = CInt(InputBox("Seriennummerlänge eingeben (0=deaktiviert, 3=001, 4=0001 usw)"))
    curSerial = 0
    bMark = True
  End If
  If bMark Then
    System.IO.File.WriteAllText(sndir+"\"+FK+".txt", CStr(zeroLen)+vbNewLine+CStr(curSerial+1)+vbNewLine)
  End If
  If zeroLen.Equals(0) Then
    Return CStr(curSerial)
  Else
    Return Right("00000000000000000000000000"+CStr(curSerial), zeroLen)
  End If
End Function

Public Function LC_Formatter(ByRef sFormat As String, ByRef sString As String, ByRef sDefault As String, ByRef bLaser As Boolean) As String Handles LC.Formatter
  If sString.Equals("MATERIAL") Then
    Return MATERIAL
  End If
  If sString.Equals("FK") Then
    Return FK
  End If

  If sString.Equals("CUSTOMSN") Then
    Return getSerial(bLaser)
  End If
End Function

Sub wsWait()
  While WORKSTATION.GetMoveState(True,True,True) Or XY_TABLE.GetMoveState(True,True,True)
    Wait 0.1
    DoEvents
  End While
End Sub
