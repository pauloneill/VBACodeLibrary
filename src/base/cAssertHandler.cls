VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAssertHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Class:        cAssertHandler
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
'
' Public Class Level Constants
'
'
' Public Class Level Properties
'
'
' Private Class Level Constants
'
'
' Private Class Level Variables
'

'
' Function:     handleAssertion
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
Public Function handleAssertion(ByVal message As String) As Long
    Dim iResponse As Integer
#If InAccess Then
    ' Show it to the user (MDBs only)
    If StrComp(Right(CurrentDb.Name, 4), ".mdb", vbTextCompare) = 0 Then
        ' Open database - assert now
        If MsgBox(message & vbCrLf & vbCrLf & "Do you wish to debug it now?", _
                  vbYesNo + vbQuestion) = vbYes Then
            iResponse = Assert.assertionStop
        Else
            iResponse = Assert.assertionContinue
        End If
        
    ElseIf StrComp(Right(CurrentDb.Name, 4), ".mde", vbTextCompare) = 0 Then
        ' Compiled database - ignore
        iResponse = Assert.assertionContinue
    Else
        ' Unknown database type - ignore it
        iResponse = Assert.assertionContinue
    End If
#Else
    ' Not Access (Excel/Word): always offer to debug
    If MsgBox(message & vbCrLf & vbCrLf & "Do you wish to debug it now?", _
              vbYesNo + vbQuestion) = vbYes Then
        iResponse = Assert.assertionStop
    Else
        iResponse = Assert.assertionContinue
    End If
#End If
    
    handleAssertion = iResponse
End Function
