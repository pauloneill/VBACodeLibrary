Attribute VB_Name = "Globals"
'
' Function:     globals
'
' Description:
'
' Assumptions:
'
' Limitations:
'
' Keywords:     Module
'
' Enhancements: 1.
'
' Author:       Paul O'Neill
'
' Written:      13 September, 2002
'
' Revision History
'
'   Date    Who     Reason
' 12Mar11   PON     Removed global Internet object.
' 10Nov04   PON     Removed constant "errCancelled" - already in modUserInput.
' 30Aug04   PON     Added global Internet object.
' 03Nov26   Claude  Host-neutral: dropped Option Compare Database (Access-only).
' 17Aug04   PON     Added global goLogger.
' 19Jul04   PON     Added global Desktop variable.
' 30Apr03   PON     Added constant "errCancelled".
' 01Nov02   PON/BVV Added constants for set names and source names.
' 13Sep02   PON     Initial coding.
'
Option Explicit             ' Declare EVERYTHING!

' Public Class Level Constants
'
'Public Const errCancelled = -32768

' Public Class Level Variables
'
Public Assert As New CAssert
'Public goLogger As New cLogger
'Public Internet As New cInternetSupport

' Private Class Level Constants
'

' Private Class Level Variables
'
