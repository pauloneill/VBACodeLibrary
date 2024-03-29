VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBusyPointer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
' Class:        CBusyPointer
'
' Description:  Displays a busy pointer for the lifetime
'               of this object.
'
' Compatability:
'
' Dependancies:
'
' Assumptions:  1.  Destroying the object will invoke
'                   the Terminate method.
'
' Limitations:  1.  There is no way to "suspend" the
'                   hourglass other than by destroying,
'                   and then recreating the object.
'
' Keywords:     Class
'
' Author:       Paul O'Neill
'
' Written:      20 May, 1998
'
' Revision History
'
'   Date    Who     Reason
' 20May98   PON     Initial coding.
'
Option Compare Database     ' Use database order for string comparisons
Option Explicit             ' Declare EVERYTHING!
Option Base 1               ' Arrays start at 1
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
' Function:     Show
'
' Description:  Displays a busy pointer in place of the
'               default mouse pointer.
'
' Parameters:   None.
'
' Dependancies:
'
' Assumptions:  1.  Busy pointer will be hidden by
'                   destroying this object.
'
' Limitations:
'
' Keywords:     Method
'
' Enhancements: 1.
'
' Author:       Paul O'Neill
'
' Written:      20 May, 1998
'
' Revision History
'
'   Date    Who     Reason
' 20May98   PON     Initial coding.
'
Public Sub Show()
    DoCmd.hourGlass True
End Sub

'
' Function:     Terminate (Destructor)
'
' Description:  Destroys any child objects
'               used by objects of this class.
'
' Assumptions:
'
' Limitations:
'
' Keywords:     Destructor
'
' Enhancements: 1.
'
' Author:       Paul O'Neill
'
' Written:      20 May, 1998
'
' Revision History
'
'   Date    Who     Reason
' 20May98   PON     Initial coding.
'
Private Sub Class_Terminate()
    DoCmd.hourGlass False
End Sub
