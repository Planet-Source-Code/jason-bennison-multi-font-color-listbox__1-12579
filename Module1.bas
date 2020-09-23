Attribute VB_Name = "Module1"
Option Explicit
'setup a mini database engine, so we can put the content to the hard drive in some sort of order.

Public MyDB As NameInfo
Type NameInfo
    Names As String * 42
    Ages As Integer
End Type
Public Filenum As Integer
Public RecordLen As Long
Public Currentrecord As Long
Public LastRecord As Long
