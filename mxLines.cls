VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mxLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'**************************************************************************************************
'  Copyright © 2003, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************
Implements mxCollEx
Dim m_colLine As mxCollEx
Dim m_Selected As Variant

'**************************************************************************************************
' mxLines Class Methods
'**************************************************************************************************
Function LineAdd(mxNewLine As mxLine, sKey As String) As mxLine
     Dim mxNew As mxLine
     Set mxNew = mxNewLine
     ' add the line to the line collection
     m_colLine.Add mxNew, sKey
End Function ' LineAdd

Function LineCount() As Long
   LineCount = m_colLine.Count
End Function ' LineCount

Function LineItem(ByVal Index As Variant) As mxLine
Attribute LineItem.VB_UserMemId = 0
   Set LineItem = m_colLine.Item(Index)
   m_Selected = Index
End Function ' LineItem

Sub LineRemove(ByVal Index As Variant)
   m_colLine.Remove Index
End Sub ' LineRemove

Sub LinesClear()
     m_colLine.Clear
End Sub ' LinesClear

Property Get Selected() As Long
     Selected = m_Selected
End Property ' Selected

'**************************************************************************************************
' mxLines - IUnknown_NewEnum gives us access to For Each...Next enumeration
'**************************************************************************************************
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = m_colLine.NewEnum
End Function ' NewEnum

'**************************************************************************************************
' mxLines - Intrinsic Methods
'**************************************************************************************************
Private Sub Class_Initialize()
     Set m_colLine = New mxCollEx
End Sub ' Class_Initialize

Private Sub Class_Terminate()
     m_colLine.Clear
     Set m_colLine = Nothing
End Sub ' Class_Terminate

'**************************************************************************************************
' mxLines - Implemented mxCollEx Collection Interface
'**************************************************************************************************
Private Sub mxCollEx_Add(newObj As Object, Optional Key As Variant)
     ' Implemented Add method
End Sub

Private Sub mxCollEx_Clear()
     ' Implemented Clear method
End Sub

Private Property Get mxCollEx_Count() As Long
     ' Implemented Count property
End Property

Private Property Get mxCollEx_Item(Index As Variant) As Variant
     ' Implemented Item  method
End Property

Private Function mxCollEx_NewEnum() As stdole.IUnknown
     ' Implemented Enum method
End Function

Private Sub mxCollEx_Remove(Index As Variant)
     ' Implemented Remove method
End Sub
