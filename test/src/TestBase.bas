Attribute VB_Name = "TestBase"
' 03Nov26   Claude  Initial smoke tests for the base module.
'
Option Explicit

' Run from the Immediate window:  RunBaseTests
Public Sub RunBaseTests()
    Debug.Print "--- base tests ---"
    TestAssertFiresAndIsCaptured
    TestAssertTrueDoesNotFire
    Debug.Print "--- done ---"
End Sub

Private Sub TestAssertFiresAndIsCaptured()
    Dim handler As New cSilentAssertionHandler
    Dim asInterface As cAssertHandler       ' addHandler takes the interface type ByRef
    Set asInterface = handler
    Assert.addHandler asInterface
    Assert.Assert False, "expected"
    Assert.removeHandler asInterface
    Report "failed assertion is captured", (handler.count = 1)
End Sub

Private Sub TestAssertTrueDoesNotFire()
    Dim handler As New cSilentAssertionHandler
    Dim asInterface As cAssertHandler
    Set asInterface = handler
    Assert.addHandler asInterface
    Assert.Assert True, "not expected"
    Assert.removeHandler asInterface
    Report "true assertion does not fire", (handler.count = 0)
End Sub

Private Sub Report(ByVal testName As String, ByVal passed As Boolean)
    Debug.Print IIf(passed, "PASS  ", "FAIL  ") & testName
End Sub
