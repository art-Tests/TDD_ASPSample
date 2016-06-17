<%
Option Explicit
%>
<!-- #include virtual="/include/ASPUnitRunner.asp"-->
<!-- #include file="CartTest.asp"-->
<%
Dim oRunner
Set oRunner = New UnitRunner
oRunner.AddTestContainer New CartTest
oRunner.Display()

%>
