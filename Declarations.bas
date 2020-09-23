Attribute VB_Name = "Declarations"
Option Explicit
Option Compare Text
Option Base 1

Public Repeat As Boolean
Public Wait As Boolean
Public MoreThan As Boolean
Public HideUnhide As Boolean
Public Mode As Boolean
Public Generating As Boolean

Public Process As Integer
Public LoopTimes As Integer
Public X As Integer

Public Key As String
Public TextString As String

Public Declare Sub Sleep Lib "kernel32" (ByVal lpMilliSeconds As Long)
