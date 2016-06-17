<%
Class TestError

	Private m_oTestCase
	Private m_lErrNumber
	Private m_sSource
	Private m_sDescription

	Public Sub Initialize(oTestCase, lErrNumber, sSource, sDescription)
		Set m_oTestCase = oTestCase
		m_lErrNumber = lErrNumber
		m_sSource = sSource
		m_sDescription = sDescription
	End Sub

	Public Property Get TestCase
		Set TestCase = m_oTestCase
	End Property

	Public Property Get ErrNumber
		ErrNumber = m_lErrNumber
	End Property

	Public Property Get Source
		Source = m_sSource
	End Property

	Public Property Get Description
		Description = m_sDescription
	End Property

End Class
%>