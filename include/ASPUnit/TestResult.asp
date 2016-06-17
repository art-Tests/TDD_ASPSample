<%
Class TestResult

	Private m_dicErrors
	Private m_dicFailures
	Private m_dicSuccesses
	Private m_dicObservers
	Private m_iRunTests
	Private m_oCurrentTestCase

	Private Sub Class_Initialize
		Set m_dicErrors = Server.CreateObject("Scripting.Dictionary")
		Set m_dicFailures = Server.CreateObject("Scripting.Dictionary")
        Set m_dicSuccesses = Server.CreateObject("Scripting.Dictionary")
		Set m_dicObservers = Server.CreateObject("Scripting.Dictionary")
        m_iRunTests = 0		
	End Sub

	Private Sub Class_Terminate
		Set m_dicErrors = Nothing
		Set m_dicFailures = Nothing
        Set m_dicSuccesses = Nothing
		Set m_dicObservers = Nothing
		Set m_oCurrentTestCase = Nothing
	End Sub

	Public Property Get Errors()
		Set Errors = m_dicErrors
	End Property

	Public Property Get Failures()
		Set Failures = m_dicFailures
	End Property

    Public Property Get Successes()
        Set Successes = m_dicSuccesses
    End Property

	Public Property Get RunTests()
		RunTests = m_iRunTests
	End Property

	Public Sub StartTest(oTestCase)
		Set m_oCurrentTestCase = oTestCase

		Dim oObserver
		For Each oObserver In m_dicObservers.Items
			oObserver.OnStartTest
		Next
	End Sub

	Public Sub EndTest()
		m_iRunTests = m_iRunTests + 1

		Dim oObserver
		For Each oObserver In m_dicObservers.Items
			oObserver.OnEndTest
		Next
	End Sub

	Public Sub AddObserver(oObserver)
		m_dicObservers.Add oOserver, oObserver
	End Sub

	Public Function AddError(oError)
		Dim oTestError
		Set oTestError = New TestError
		oTestError.Initialize m_oCurrentTestCase, oError.Number, oError.Source, oError.Description
		m_dicErrors.Add oTestError, oTestError

		Dim oObserver
		For Each oObserver In m_dicObservers.Items
			oObserver.OnError
		Next

		Set AddError = oTestError
	End Function

	Public Function AddFailure(sMessage)
		Dim oTestError
		Set oTestError = New TestError
		oTestError.Initialize m_oCurrentTestCase, 0, " ", sMessage
		m_dicFailures.Add oTestError, oTestError

		Dim oObserver
		For Each oObserver In m_dicObservers.Items
			oObserver.OnFailure
		Next

		Set AddFailure = oTestError
	End Function

    Public Function AddSuccess
		Dim oTestError
		Set oTestError = New TestError
		oTestError.Initialize m_oCurrentTestCase, 0, " ", "測試成功！沒有失敗發生！"
		m_dicSuccesses.Add oTestError, oTestError

		Dim oObserver
		For Each oObserver In m_dicObservers.Items
			oObserver.OnSuccess
		Next
    End Function

	Public Sub Assert(bCondition, sMessage)
	    If Not bCondition Then
		    AddFailure sMessage
		End If
	End Sub

	Public Sub AssertEquals(vExpected, vActual, sMessage)
		If vExpected <> vActual Then
			AddFailure NotEqualsMessage(sMessage, vExpected, vActual)
		End	If
	End Sub

	' Build a message about a failed equality check
	Function NotEqualsMessage(sMessage, vExpected, vActual)
		NotEqualsMessage = sMessage & "<br /> - 預期結果為 " & CStr(vExpected) & " 但實際結果為  " & CStr(vActual) & " ，應該要相等。"
	End Function

	Public Sub AssertNotEquals(vExpected, vActual, sMessage)
		If vExpected = vActual Then
			AddFailure sMessage & "<br /> - 預期結果為 " & CStr(vExpected) & " 但實際結果為 " & CStr(vActual) & " ，應該要不相等。"
		End	If
	End Sub

	Public Sub AssertExists(vVariable, sMessage)
		If IsObject(vVariable) Then
			If (vVariable Is Nothing) Then
				AddFailure sMessage & "<br /> - 變數的型態為 " & TypeName(vVariable)
			End If
		ElseIf IsEmpty(vVariable) Then
			AddFailure sMessage & "<br /> - 變數的型態為 " & TypeName(vVariable) & " (未初始化)."
		End If
	End Sub

End Class
%>