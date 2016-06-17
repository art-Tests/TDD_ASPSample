<%
'********************************************************************
' Name: ASPUnitRunner.asp
'
' Purpose: Contains the UnitRunner class which is used to render the unit testing UI
'********************************************************************

'********************************************************************
' Include Files
'********************************************************************
%>
<!-- #include file="ASPUnit.asp"-->
<%

Const ALL_TESTCONTAINERS = "所有測試容器"
Const ALL_TESTCASES = "所有測試案例"

Const FRAME_PARAMETER = "UnitRunner"
Const FRAME_SELECTOR = "selector"
Const FRAME_RESULTS = "results"

Const STYLESHEET = "../include/ASPUnit.css"
Const SCRIPTFILE = "../include/ASPUnitRunner.js"

Class UnitRunner

	Private m_dicTestContainer

	Private Sub Class_Initialize()
		Set m_dicTestContainer = CreateObject("Scripting.Dictionary")
	End Sub

	Public Sub AddTestContainer(oTestContainer)
		m_dicTestContainer.Add TypeName(oTestContainer), oTestContainer
	End Sub

	Public Function Display()
		If (Request.QueryString(FRAME_PARAMETER) = FRAME_SELECTOR) Then
			DisplaySelector
		ElseIf (Request.QueryString(FRAME_PARAMETER) = FRAME_RESULTS) Then
			DisplayResults
		Else
			ShowFrameSet
		End if
	End Function

'********************************************************************
' Frameset
'********************************************************************
	Private Function ShowFrameSet()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ASPUnit Test Runner</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
</head>
<frameset rows="30, *">
<noframes>
<body>
<p>抱歉，您的瀏覽器無法支援頁框。</p>
</body>
</noframes>
<frame name="<% = FRAME_SELECTOR %>" src="<% = GetSelectorFrameSrc %>" marginwidth="0" marginheight="0" scrolling="auto" frameborder="0" noresize="noresize" />
<frame name="<% = FRAME_RESULTS %>" src="<% = GetResultsFrameSrc %>" marginwidth="0" marginheight="0" scrolling="auto" frameborder="0" noresize="noresize" />
</frameset>
</html>
<%
	End Function

	Private Function GetSelectorFrameSrc()
		GetSelectorFrameSrc = Request.ServerVariables("SCRIPT_NAME") & "?" & FRAME_PARAMETER & "=" & FRAME_SELECTOR
	End Function

	Private Function GetResultsFrameSrc()
		GetResultsFrameSrc = Request.ServerVariables("SCRIPT_NAME") & "?" & FRAME_PARAMETER & "=" & FRAME_RESULTS
	End Function

'********************************************************************
' Selector Frame
'********************************************************************
	Private Function DisplaySelector()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>控制台</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<link rel="stylesheet" href="<% = STYLESHEET %>" media="screen" type="text/css" />
<script type="text/javascript" src="<% = SCRIPTFILE %>"></script>
</head>
<body>
<form name="frmSelector" action="<% = GetResultsFrameSrc %>" target="<% = FRAME_RESULTS %>" method="post">
<table width="80%" align="center">
<tr>
<td align="right" nowrap="nowrap">欲執行的測試 :</td>
<td>
<select name="cboTestContainers" onchange="ComboBoxUpdate('<% = GetSelectorFrameSrc %>', '<% = FRAME_SELECTOR %>');">
<option><% = ALL_TESTCONTAINERS %></option>
<% AddTestContainers %>
</select>
</td>
<td align="right" nowrap="nowrap">欲測試的案例 :</td>
<td>
<select name="cboTestCases">
<option><% = ALL_TESTCASES %></option>
<% AddTestMethods %>
</select>
</td>
<td nowrap="nowrap" colspan="2">
<input type="checkbox" name="chkShowSuccess" id="chkShowSuccess" checked="checked" /><label for="chkShowSuccess">顯示通過的測試</label>
</td>
<td>
<input type="submit" name="cmdRun" value="執行測試" />
</td>
</tr>
</table>
</form>
</body>
</html>
<%
	End Function

	Private Function AddTestContainers()
		Dim oTestContainer, sTestContainer
		For Each oTestContainer In m_dicTestContainer.Items()
			sTestContainer = TypeName(oTestContainer)
			If (sTestContainer = Request.Form("cboTestContainers")) Then
				Response.Write("<option selected=""selected"">" & sTestContainer & "</option>")
			Else
				Response.Write("<option>" & sTestContainer & "</option>")
			End If
		Next
	End Function

	Private Function AddTestMethods()
		Dim oTestContainer, sContainer, sTestMethod

		If (Request.Form("cboTestContainers") <> ALL_TESTCONTAINERS And Request.Form("cboTestContainers") <> "") Then
			sContainer = CStr(Request.Form("cboTestContainers"))
			Set oTestContainer = m_dicTestContainer.Item(sContainer)
			For Each sTestMethod In oTestContainer.TestCaseNames()
				Response.Write("<option>" & sTestMethod & "</option>" & vbCrLf)
			Next
		End If
	End Function

	Private Function TestName(oResult)
		If (oResult.TestCase Is Nothing) Then
			TestName = ""
		Else
			TestName = TypeName(oResult.TestCase.TestContainer) & "." & oResult.TestCase.TestMethod
		End If
	End Function

'********************************************************************
' Results Frame
'********************************************************************
	Private Function DisplayResults()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>執行結果</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<link rel="stylesheet" href="<% = STYLESHEET %>" media="screen" type="text/css" />
</head>
<body>
<%
		Dim oTestResult, oTestSuite
		Set oTestResult = New TestResult

		' Create TestSuite
		Set oTestSuite = BuildTestSuite()

		' Run Tests
		oTestSuite.Run oTestResult

		' Display Results
		DisplayResultsTable oTestResult
%>
<p align="center">
<a href="http://validator.w3.org/check?uri=referer"><img src="http://www.w3.org/Icons/valid-xhtml10" alt="Valid XHTML 1.0 Transitional" height="31" width="88" border="0" /></a>
</p>
</body>
</html>
<%
	End Function

	Private Function BuildTestSuite()

		Dim oTestSuite, oTestContainer, sContainer
		Set oTestSuite = New TestSuite

		If (Request.Form("cmdRun") <> "") Then
			If (Request.Form("cboTestContainers") = ALL_TESTCONTAINERS) Then
				For Each oTestContainer In m_dicTestContainer.Items()
					If Not(oTestContainer Is Nothing) Then
						oTestSuite.AddAllTestCases oTestContainer
					End If
				Next
			Else
				sContainer = CStr(Request.Form("cboTestContainers"))
				Set oTestContainer = m_dicTestContainer.Item(sContainer)

				Dim sTestMethod
				sTestMethod = Request.Form("cboTestCases")

				If (sTestMethod = ALL_TESTCASES) Then
					oTestSuite.AddAllTestCases oTestContainer
				Else
					oTestSuite.AddTestCase oTestContainer, sTestMethod
				End If
			End If
		End If

		Set BuildTestSuite = oTestSuite
	End Function

	Private Function DisplayResultsTable(oTestResult)
%>
<table border="1" width="80%" align="center" id="test-result">
<tr><th width="15%">型態 (Type)</th><th width="20%">測試案例 (Test)</th><th width="70%">描述 (Description)</th></tr>
<%
		If Not(oTestResult Is Nothing) Then
			Dim oResult
			If (Request.Form("chkShowSuccess") <> "") Then
	            For Each oResult in oTestResult.Successes
					Response.Write("<tr class=""success""><td>成功 (Success)</td><td>" & TestName(oResult) & "</td><td>" & oResult.Source & oResult.Description & "</td></tr>" & vbCrLf)
	            Next
	        End If

			For Each oResult In oTestResult.Errors
				Response.Write("<tr class=""error""><td>錯誤 (Error)</td><td>" & TestName(oResult) & "</td><td>" & oResult.Source & " (" & Trim(oResult.ErrNumber) & "): " & oResult.Description & "</td></tr>" & vbCrLf)
			Next

			For Each oResult In oTestResult.Failures
				Response.Write("<tr class=""warning""><td>失敗 (Failure)</td><td>" & TestName(oResult) & "</td><td>" & oResult.Description & "</td></tr>" & vbCrLf)
			Next

			Response.Write("<tr><td align=""center"" colspan=""3"">" & "共測試 " & oTestResult.RunTests & " 個案例，有 " & oTestResult.Errors.Count & " 個錯誤 (Errors) 及 " & oTestResult.Failures.Count & " 個失敗 (Failures)</td></tr>" & vbCrLf)
		End If
%>
</table>
<%
	End Function

	Public Sub OnStartTest()

	End Sub

	Public Sub OnEndTest()

	End Sub

	Public Sub OnError()

	End Sub

	Public Sub OnFailure()

	End Sub

    Public Sub OnSuccess()

    End Sub
End Class
%>

