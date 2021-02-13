VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class Variables
Private mValue As String
Private mName As String

' Retrieve Value
Public Property Get Value() As String
    Value = mValue
End Property

' Set a new Value
Public Property Let Value(ByVal iNewValue As String)
    mValue = iNewValue
End Property

' Retrieve Name
Public Property Get Name() As String
    Name = mName
End Property

' Set a new Name
Public Property Let Name(ByVal strNewName As String)
    mName = strNewName
End Property

