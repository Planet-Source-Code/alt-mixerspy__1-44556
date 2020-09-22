VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mxCollEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'**************************************************************************************************
'  Copyright © 2003, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************

'**************************************************************************************************
' mxCollEx - Object Variable To Wrap Collection Object
'**************************************************************************************************
Private m_Col As Collection

'**************************************************************************************************
' mxCollEx Property Read-Only properties
'**************************************************************************************************
Public Property Get Count() As Long
     ' exit if m_col is not set
     If m_Col Is Nothing Then Exit Property
     Count = m_Col.Count
End Property ' Get Count

Public Property Get Item(Index As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    On Error Resume Next
    Set Item = m_Col.Item(Index)
End Property ' Get Item

'**************************************************************************************************
' mxCollEx Methods
'**************************************************************************************************
Public Sub Add(newObj As Object, Optional Key As Variant)
     ' What to do if missing key
     If IsMissing(Key) Then
          m_Col.Add newObj
     Else
          m_Col.Add newObj, CStr(Key)
     End If
End Sub ' Add

Public Sub Clear()
     Dim iLoop As Integer
     ' Shouldn't happen, but exit in case.
     If m_Col Is Nothing Then Exit Sub
     For iLoop = m_Col.Count To 1 Step -1
          m_Col.Remove iLoop
     Next
End Sub ' Clear

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    ' facilitates "For..Each" enumeration functionality
    Set NewEnum = m_Col.[_NewEnum]
End Function ' NewEnum

Public Sub Remove(Index As Variant)
    On Error Resume Next
    ' Remove item passed in arg
    m_Col.Remove Index
End Sub ' Remove

Private Sub Class_Initialize()
     ' Create the new collection
     Set m_Col = New Collection
End Sub ' Class_Initialize

Private Sub Class_Terminate()
     ' Empty collection
     Clear
     ' kill it
     Set m_Col = Nothing
End Sub ' Class_Terminate
