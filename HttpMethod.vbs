Function HTTP_GET(URL)
	Dim XML
	Set XML = CreateObject("WinHttp.WinHttpRequest.5.1")
	With XML
		.Open "GET",URL + "?t=" + CStr(Rnd()),True'�����ʱ��������棬���URL�в����Ͱ�?��Ϊ&
		.SetTimeouts 50000, 50000, 50000, 50000'��ʱ
		.Send
		.WaitForResponse
		If XML.Status = 200 Then'�ɹ���ȡҳ��
			HTTP_GET = GB2312ToUnicode(.ResponseBody)
		Else
			MsgBox "Http�������:" & .Status, 16, "��ʾ"
		End If
	End With
End Function

Function HTTP_POST(URL,data)
	Dim XML
	Set XML = CreateObject("WinHttp.WinHttpRequest.5.1")
	With XML
		.Open "POST",URL ,True
		.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"'����HTTPͷ��Ϣ
		.SetTimeouts 50000, 50000, 50000, 50000'��ʱ
		.Send(data)
		.WaitForResponse
		If XML.Status = 200 Then'�ɹ���ȡҳ��
			HTTP_POST = GB2312ToUnicode(.ResponseBody)
		Else
			MsgBox "Http�������:" & .Status, 16, "��ʾ"
		End If
	End With
End Function

Function GB2312ToUnicode(str)
	With CreateObject("adodb.stream")
		.Type = 1 : .Open
		.Write str : .Position = 0
		.Type = 2 : .Charset = "gb2312"
		GB2312ToUnicode = .ReadText : .Close
	End With
End Function
