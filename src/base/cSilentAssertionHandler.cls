VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSilentAssertionHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Class:        cSilentAssertHandler
'
' Description:
'
' Compatability:
'
' Dependancies:
'
' Assumptions:  1.
'
' Limitations:
'
' Keywords:     Class
'
' Author:       Paul O'Neill, o'neill software Pty Ltd
'
' Written:      06 December, 2005
'
' Revision History
'
'   Date    Who     Reason
' 03Nov26   Claude  Host-neutral: dropped Option Compare Database; Access-only code under #If InAccess.
' 06Dec05   PON     Initial coding.
'
Option Explicit             ' Declare EVERYTHING!

Implements cAssertHandler

' Public Class Level Constants
'

' Public Class Level Properties
'

' Private Class Level Constants
'

' Private Class Level Variables
'
Private mlAssertions As Long
Private mtAssertions As String

'
' Function:     count (Property Get)
'
' Description:  Gets the count property
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
Public Property Get count()
    ' Return the current value of count
    count = mlAssertions
End Property

'
' Function:     assertions (Property Get)
'
' Description:  Gets the assertions property
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
Public Property Get assertions()
    ' Return the current value of assertions
    assertions = mtAssertions
End Property

'
' Function:     reset
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
Public Function reset()
    mlAssertions = 0
    mtAssertions = ""
End Function

'
' Function:     cAssertHandler_handleAssertion
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
Private Function cAssertHandler_handleAssertion(ByVal message As String) As Long
    mlAssertions = mlAssertions + 1
    mtAssertions = mtAssertions + message + vbCrLf
End Function

'
' Function:     Class_Initialize
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
Private Sub Class_Initialize()
    reset
End Sub
