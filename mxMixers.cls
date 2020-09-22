VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mxMixers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'**************************************************************************************************
'  Copyright © 2003, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************
Implements mxCollEx
Dim m_colMxx As mxCollEx
Dim m_Selected As Variant

'**************************************************************************************************
' mxMixers Class Methods
'**************************************************************************************************
Function MixerAdd(mxNewMixer As mxMixer, sKey As String) As mxMixer
     Dim mxNew As mxMixer
     Set mxNew = mxNewMixer
     ' add the mixer to the collection
     m_colMxx.Add mxNew, sKey
End Function ' MixerAdd

Function MixerCount() As Long
   MixerCount = m_colMxx.Count
End Function ' MixerCount

Function MixerItem(ByVal Index As Variant) As mxMixer
Attribute MixerItem.VB_UserMemId = 0
   Set MixerItem = m_colMxx.Item(Index)
   m_Selected = Index
End Function ' MixerItem

Sub MixerRemove(ByVal Index As Variant)
   m_colMxx.Remove Index
End Sub ' MixersRemove

Sub MixersClear()
     m_colMxx.Clear
End Sub ' MixersClear

Property Get Selected() As Long
     Selected = m_Selected
End Property ' Get Selected

'**************************************************************************************************
' mxMixers - IUnknown_NewEnum gives us access to For Each...Next enumeration
'**************************************************************************************************
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = m_colMxx.NewEnum
End Function ' NewEnum

'**************************************************************************************************
' mxMixers - Intrinsic Methods
'**************************************************************************************************
Private Sub Class_Initialize()
     Set m_colMxx = New mxCollEx
End Sub ' Class_Initialize

Private Sub Class_Terminate()
     m_colMxx.Clear
     Set m_colMxx = Nothing
End Sub ' Class_Terminate

'**************************************************************************************************
' mxMixers - Implemented mxCollEx Collection Interface
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
