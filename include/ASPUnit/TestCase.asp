<%
Class TestCase
	Private m_oTestContainer
	Private m_sTestMethod

	Public Property Get TestContainer()
		Set TestContainer = m_oTestContainer
	End Property

	Public Property Set TestContainer(oTestContainer)
		Set m_oTestContainer = oTestContainer
	End Property

	Public Property Get TestMethod()
		TestMethod = m_sTestMethod
	End Property

	Public Property Let TestMethod(sTestMethod)
		m_sTestMethod = sTestMethod
	End Property

	Public Sub Run(oTestResult)

        Dim iOldFailureCount
        Dim iOldErrorCount

        iOldFailureCount = oTestResult.Failures.Count
        iOldErrorCount = oTestResult.Errors.Count

		On Error Resume Next
		oTestResult.StartTest Me

		m_oTestContainer.SetUp()

		If (Err.Number <> 0) Then
			oTestResult.AddError Err
		Else
			' Response.Write("m_oTestContainer." & m_sTestMethod & "(oTestResult)<br />")
			Execute("m_oTestContainer." & m_sTestMethod & "(oTestResult)")

			If (Err.Number <> 0) Then
				' Response.Write(Err.Description & "<br />")
				oTestResult.AddError Err
			End	If
		End If
		Err.Clear()

		m_oTestContainer.TearDown()

        If (Err.Number <> 0) Then
			oTestResult.AddError Err
		End If

		'Log success if no failures or errors occurred
		If oTestResult.Failures.Count = iOldFailureCount And oTestResult.Errors.Count = iOldErrorCount Then
		    oTestResult.AddSuccess
		End If
		
        oTestResult.EndTest
        
		On Error Goto 0
	End Sub

End Class
%>