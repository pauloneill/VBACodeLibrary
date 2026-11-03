VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAssert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 06Dec05   PON     Added assertionHandlers property and code to use it.
' 03Nov26   Claude  Host-neutral: dropped Option Compare Database (Access-only); Access code stays under #If InAccess.

Option Explicit

Private Enum assertionResponse
    assertQuit = -1
    assertContinue = 0
    assertStop = 1
End Enum

Private mcolHandlers As New Collection

Private moDefaultHandler As New cAssertHandler

'
' Function:     assertionQuit (Property Get)
'
' Description:  Gets the assertionQuit property
'               for this object.
'
' Parameters:   None.
'
' Returns:
'
' Assumptions:  None.
'
' Limitations:
'
' Keywords:     Property
'
' Enhancements: 1.
'
' Author:       Paul O'Neill, o'neill software Pty Ltd
'
' Written:      06 December, 2005
'
' Revision History
'
'   Date    Who     Reason
' 06Dec05   PON     Initial coding.
'
Public Property Get assertionQuit() As Long
    ' Return the current value of assertionQuit
    assertionQuit = assertionResponse.assertQuit
End Property

'
' Function:     assertionStop (Property Get)
'
' Description:  Gets the assertionStop property
'               for this object.
'
' Parameters:   None.
'
' Returns:
'
' Assumptions:  None.
'
' Limitations:
'
' Keywords:     Property
'
' Enhancements: 1.
'
' Author:       Paul O'Neill, o'neill software Pty Ltd
'
' Written:      06 December, 2005
'
' Revision History
'
'   Date    Who     Reason
' 06Dec05   PON     Initial coding.
'
Public Property Get assertionStop() As Long
    ' Return the current value of assertionStop
    assertionStop = assertionResponse.assertStop
End Property

'
' Function:     assertionContinue (Property Get)
'
' Description:  Gets the assertionContinue property
'               for this object.
'
' Parameters:   None.
'
' Returns:
'
' Assumptions:  None.
'
' Limitations:
'
' Keywords:     Property
'
' Enhancements: 1.
'
' Author:       Paul O'Neill, o'neill software Pty Ltd
'
' Written:      06 December, 2005
'
' Revision History
'
'   Date    Who     Reason
' 06Dec05   PON     Initial coding.
'
Public Property Get assertionContinue() As Long
    ' Return the current value of assertionContinue
    assertionContinue = assertionResponse.assertContinue
End Property

'
' Function:     addHandler
'
' Description:
'
' Parameters:   1.
'
' Returns:
'
' Dependancies:
'
' Assumptions:
'
' Limitations:
'
' Keywords:
'
' Enhancements: 1.
'
' Author:       Paul O'Neill, o'neill software Pty Ltd
'
' Written:      06 December, 2005
'
' Revision History
'
'   Date    Who     Reason
' 06Dec05   PON     Initial coding.
'
Public Sub addHandler(newHandler As cAssertHandler)
    mcolHandlers.Add newHandler
End Sub

'
' Function:     removeHandler
'
' Description:
'
' Parameters:   1.
'
' Returns:
'
' Dependancies:
'
' Assumptions:
'
' Limitations:
'
' Keywords:
'
' Enhancements: 1.
'
' Author:       Paul O'Neill, o'neill software Pty Ltd
'
' Written:      06 December, 2005
'
' Revision History
'
'   Date    Who     Reason
' 06Dec05   PON     Initial coding.
'
Public Sub removeHandler(oldHandler As cAssertHandler)
    Dim lIndex As Long
    For lIndex = 1 To mcolHandlers.count
        If mcolHandlers(lIndex) Is oldHandler Then
            mcolHandlers.Remove lIndex
            Exit Sub
        End If
    Next lIndex
    
    Debug.Assert False
End Sub


Public Sub Precondition(ByVal Condition As Boolean, _
                        ByVal message As String, _
                        Optional ByVal Location As String = "")
    Const AssertionType = "PreCondition"
    
    ' What action do we take if the assertion failed?
    DisplayAssertion AssertionType, Condition, message, Location
End Sub

Public Sub Assert(ByVal Condition As Boolean, _
                  ByVal message As String, _
                  Optional ByVal Location As String = "")
    Const AssertionType = "Assert"
    
    ' What action do we take if the assertion failed?
    DisplayAssertion AssertionType, Condition, message, Location
End Sub

Public Sub PostCondition(ByVal Condition As Boolean, _
                         ByVal message As String, _
                         Optional ByVal Location As String = "")
    Const AssertionType = "PostCondition"
    
    ' What action do we take if the assertion failed?
    DisplayAssertion AssertionType, Condition, message, Location
End Sub

Private Sub DisplayAssertion(ByVal AssertionType As String, _
                             ByVal Condition As Boolean, _
                             ByVal message As String, _
                             Optional ByVal Location As String = "")

    If Condition Then
        ' Condition is true - assertion has been satisfied - do nothing
    Else
        ' Construct an assertion  message
        Dim tMsg As String
        tMsg = AssertionType & " [" & message & "] fired"
        If Location <> "" Then tMsg = tMsg & " at " & Location
        
        ' ALWAYS display it in the Immediate/Debug window
        Debug.Print tMsg
        
        Dim iResponse As Long
        If mcolHandlers.count = 0 Then
            'iResponse = moDefaultHandler.handleAssertion(tMsg)
#If InAccess Then
            If StrComp(Right(CurrentDb.Name, 4), ".mdb", vbTextCompare) = 0 Then
                ' Open database - assert now
                If MsgBox(tMsg & vbCrLf & vbCrLf & "Do you wish to debug it now?", _
                          vbYesNo + vbQuestion) = vbYes Then
                    iResponse = assertStop
                Else
                    iResponse = assertContinue
                End If
                
            ElseIf StrComp(Right(CurrentDb.Name, 4), ".mde", vbTextCompare) = 0 Then
                ' Compiled database - ignore
                iResponse = assertContinue
            Else
                ' Unknown database type - ignore it
                iResponse = assertContinue
            End If
#Else
            If MsgBox(tMsg & vbCrLf & vbCrLf & "Do you wish to debug it now?", _
                      vbYesNo + vbQuestion) = vbYes Then
                iResponse = assertStop
            Else
                iResponse = assertContinue
            End If
#End If
        Else
            ' ToDo: invoke _all_ handlers
            Debug.Assert mcolHandlers.count = 1
            iResponse = mcolHandlers(1).handleAssertion(tMsg)
        End If
        
        ' Well, what are we doing?
        Select Case iResponse
        Case assertionResponse.assertQuit
            ' BAIL! BAIL! BAIL!
            End
            
        Case assertionResponse.assertStop
            ' Interrupt the program flow
            Stop
            
        Case assertionResponse.assertContinue
            ' Do nothing
        Case Else
            ' We shouldn't get here - make a noise so thay come and investigate
            Beep
        End Select
    End If
End Sub

