On Error Resume Next

Class Iper
	Sub waitToConnect()
		isConnected = False
		While isConnected = False
			ip = "www.baidu.com"
			Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address = '" & ip & "'")
			For Each objStatus in objPing
				If objStatus.ResponseTime > 0 Then
					isConnected = True
					'Msgbox "Connected"
					exit For
				Else 
					WScript.sleep 20000
				End If
			Next
		Wend
	End Sub

	Function getIp()
		Set objXML = WScript.GetObject("http://www.ip138.com/ip2city.asp")
		While objXML.readyState = "loading"
			WScript.Sleep 100
		Wend
		Set coll = objXML.getElementsByTagName("body")
		'WScript.Echo coll(0).innertext
		ip = Split(coll(0).innertext, "£º")(1)
		Set objXML = Nothing
		getIp = ip
	End Function

	Function isIpSame(ipNew)
		Const iplog = "C:\Users\Wayne\AppData\Local\iplog.txt"
		Const ForReading = 1
		Const ForWriting = 2
		Const ForAppending = 8
		isEqual = False
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		If objFSO.FileExists(iplog) Then
			Set objFile = objFSO.OpenTextFile(iplog, ForReading)
			ipOld = Replace(objFile.ReadAll, vbNewLine, "") 'Mark, read a line with '\n', delete it
			objFile.Close
			'MsgBox StrComp(Trim(CStr(ipNew)), Trim(CStr(ipOld)), 1) 
			If Trim(CStr(ipNew)) = Trim(CStr(ipOld)) Then
				isEqual = True
				'MsgBox "Equal"
			Else 
				Set objFile = objFSO.OpenTextFile(iplog, ForWriting)
				objFile.WriteLine(ipNew)
				objFile.Close
			End If 
		Else
			'Wscript.Echo "File does not exist."
			Set objFile = objFSO.CreateTextFile(iplog, True)
			objFile.WriteLine(ipNew)
			objFile.Close
			isEqual = False
		End If
		Set objFSO = Nothing
		'MsgBox isEqual
		isIpSame = isEqual
	End Function

	Sub sendEmail(ip)
		Const Email_From = "zhengwei_hit@yeah.net"
		Const Password = "0987654321"
		Const Email_To = "zwee.cn@139.com"
		Set CDO = CreateObject("CDO.Message")
		CDO.Subject = "IP of My Computer"
		CDO.From = Email_From
		CDO.To = Email_To
		msg = CStr(Now()) + " @ " + ip ' Here
		'MsgBox msg
		CDO.TextBody = msg
		'cdo.AddAttachment "C:\hello.txt"
		Const schema = "http://schemas.microsoft.com/cdo/configuration/"
		With CDO.Configuration.Fields
			.Item(schema & "sendusing") = 2
			.Item(schema & "smtpserver") = "smtp.yeah.net"
			.Item(schema & "smtpauthenticate") = 1
			.Item(schema & "sendusername") = Email_From
			.Item(schema & "sendpassword") = Password
			.Item(schema & "smtpserverport") = 25
			.Item(schema & "smtpusessl") = True
			.Item(schema & "smtpconnectiontimeout") = 60
			.Update
		End With
		CDO.Send
		Set objXML = Nothing
		'Msgbox "Message has send"
	End Sub

	Public Sub Listen()
	MsgBox "In listening."
	isFirstRun = True
		While True
			waitToConnect
			ip = getIp()
			If (Not isIpSame(ip)) Or (isFirstRun) Then
				'MsgBox "Sending email..."
				sendEmail ip
				isFirstRun = False
			End If 
			WSCript.Sleep 20000
		WEnd
	End Sub

End Class 

Set myiper = new Iper
myiper.Listen