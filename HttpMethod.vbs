Function HTTP_GET(URL)
	Dim XML
	Set XML = CreateObject("WinHttp.WinHttpRequest.5.1")
	With XML
		.Open "GET",URL + "?t=" + CStr(Rnd()),True'后面加时间戳防缓存，如果URL有参数就把?改为&
		.SetTimeouts 50000, 50000, 50000, 50000'超时
		.Send
		.WaitForResponse
		If XML.Status = 200 Then'成功获取页面
			HTTP_GET = GB2312ToUnicode(.ResponseBody)
		Else
			MsgBox "Http错误代码:" & .Status, 16, "提示"
		End If
	End With
End Function

Function HTTP_POST(URL,data)
	Dim XML
	Set XML = CreateObject("WinHttp.WinHttpRequest.5.1")
	With XML
		.Open "POST",URL ,True
		.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"'设置HTTP头信息
		.SetTimeouts 50000, 50000, 50000, 50000'超时
		.Send(data)
		.WaitForResponse
		If XML.Status = 200 Then'成功获取页面
			HTTP_POST = GB2312ToUnicode(.ResponseBody)
		Else
			MsgBox "Http错误代码:" & .Status, 16, "提示"
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
