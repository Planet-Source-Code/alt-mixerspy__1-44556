VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mxControls"
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
Dim m_colCtrl As mxCollEx
Dim m_Selected As Variant

'**************************************************************************************************
' mxLineControls Class Methods
'**************************************************************************************************
Function CtrlAdd(mxNewCtrl As mxControl, sKey As String) As mxControl
     Dim mxCtrl As mxControl
     Set mxCtrl = mxNewCtrl
     ' add the line to the line collection
     m_colCtrl.Add mxCtrl, sKey
End Function ' CtrlAdd

Function CtrlCount() As Long
   CtrlCount = m_colCtrl.Count
End Function ' CtrlCount

Function CtrlItem(ByVal Index As Variant) As mxControl
Attribute CtrlItem.VB_UserMemId = 0
   Set CtrlItem = m_colCtrl.Item(Index)
   m_Selected = Index
End Function ' CtrlItem

Sub CtrlRemove(ByVal Index As Variant)
   m_colCtrl.Remove Index
End Sub ' CtrlRemove

Sub CtrlsClear()
     m_colCtrl.Clear
End Sub ' CtrlsClear

Property Get Selected() As Long
     Selected = m_Selected
End Property ' Selected

'**************************************************************************************************
' mxLineControls - IUnknown_NewEnum gives us access to For Each...Next enumeration
'**************************************************************************************************
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
   Set NewEnum = m_colCtrl.NewEnum
End Function ' NewEnum

'**************************************************************************************************
' mxLineControls - Intrinsic Methods
'**************************************************************************************************
Private Sub Class_Initialize()
     Set m_colCtrl = New mxCollEx
End Sub ' Class_Initialize

Private Sub Class_Terminate()
     m_colCtrl.Clear
     Set m_colCtrl = Nothing
End Sub ' Class_Terminate

'**************************************************************************************************
' mxLineControls - Implemented mxCollEx Collection Interface
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


