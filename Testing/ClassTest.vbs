option explicit

!INC Local Scripts.EAConstants-VBScript
!INC EAScriptLib.VBScript-Logging

'
' Script Name: ClassTest
' Author: Philippe Back
' Purpose: Test VBScript features in the EA Scripting
' Date: 18/11/2010
'

Class TestClass

  Public m_test

  Private Sub Class_Initialize
     ' Statements go here.
	 MsgBox("Initialize")
	 LOGInfo("Initialize")
  End Sub

  Property Get TestXXX
     TestXXX = Me.m_test
  End Property

  Property Let TestXXX(test)
    Me.m_test = test
   ' Validation statements go here.
  End Property

  Sub doThis
    MsgBox("Hello")  
  End Sub
  
  Function doIt(str)
    MsgBox("Hello Function(" & str & ")")  
	m_test = "MEMBER"
    doIt=2
  End Function
End Class

sub main
	Dim oTest
	Set oTest = new TestClass
	oTest.doThis()
	Dim iVal
	iVal = oTest.doIt("Me")
	MsgBox("Doit=" & iVal)
	MsgBox("Member:" & oTest.m_test)
	oTest.TestXXX = "TEST"
	MsgBox("Property:" & oTest.TestXXX)
	
end sub

main