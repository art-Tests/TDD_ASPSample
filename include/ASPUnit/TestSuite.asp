<%
Class TestSuite
	Private m_oTestCases

	Private Sub Class_Initialize()
		Set m_oTestCases = Server.CreateObject("Scripting.Dictionary")
	End Sub

	Private Sub Class_Terminate()
		Set m_oTestCases = Nothing
	End Sub

	Public Sub AddTestCase(oTestContainer, sTestMethod)
		Dim oTestCase
		Set oTestCase = New TestCase
		Set oTestCase.TestContainer = oTestContainer
		oTestCase.TestMethod = sTestMethod

		m_oTestCases.Add oTestCase, oTestCase
	End Sub

	Public Sub AddAllTestCases(oTestContainer)
		Dim oTestCase, sTestMethod

		For Each sTestMethod In oTestContainer.TestCaseNames()
			AddTestCase oTestContainer, sTestMethod
		Next
	End Sub

	Public Function Count()
		Count = m_oTestCases.Count
	End Function

	Public Sub Run(oTestResult)
		Dim oTestCase
		For Each oTestCase In m_oTestCases.Items
			oTestCase.Run oTestResult
		Next
	End Sub
End Class
%>