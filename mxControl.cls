VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mxControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'**************************************************************************************************
'  Copyright © 2003, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************
' Memory manipulation declares
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
     pSource As Any, ByVal ByteLen As Long)
Private Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, _
     struct As Any, ByVal cb As Long)
Private Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, _
     ByVal ptr As Long, ByVal cb As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
     ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long

'**************************************************************************************************
' mxLineControl Property Variables
'**************************************************************************************************
Private m_Control As Long
Private m_ControlDesc As String
Private m_ControlID As Long
Private m_ControlName As String
Private m_ControlShortName As String
Private m_ControlType As Long
Private m_hMixer As Long
Private m_Key As String
Private m_Maximum As Long
Private m_Minimum As Long
Private m_MultipleItems As Long
Private m_Parent As String
Private m_Value As Long

'**************************************************************************************************
' mxLineControl Property Pairs
'**************************************************************************************************
Friend Property Get Control() As Long
     Control = m_Control
End Property ' Get Control

Friend Property Let Control(New_Control As Long)
     m_Control = New_Control
End Property ' Let Control

Friend Property Get ControlDesc() As String
     ControlDesc = m_ControlDesc
End Property   ' Get ControlDesc

Friend Property Let ControlDesc(New_ControlDesc As String)
     m_ControlDesc = New_ControlDesc
End Property ' Let ControlDesc

Friend Property Get ControlID() As Long
     ControlID = m_ControlID
End Property ' Get ControlID

Friend Property Let ControlID(New_ControlID As Long)
     m_ControlID = New_ControlID
End Property ' Let ControlID

Friend Property Get ControlName() As String
     ControlName = m_ControlName
End Property ' Get ControlName

Friend Property Let ControlName(New_ControlName As String)
     m_ControlName = New_ControlName
End Property ' Let ControlName

Friend Property Get ControlShortName() As String
     ControlShortName = m_ControlShortName
End Property ' Get ControlShortName

Friend Property Let ControlShortName(New_ControlShortName As String)
     m_ControlShortName = New_ControlShortName
End Property ' Let ControlShortName

Friend Property Get ControlType() As Long
     ControlType = m_ControlType
End Property ' Get ControlType

Friend Property Let ControlType(New_ControlType As Long)
     m_ControlType = New_ControlType
End Property ' Let ControlType
     
Friend Property Get hMixer() As Long
     hMixer = m_hMixer
End Property ' Get hMixer

Friend Property Let hMixer(New_hMixer As Long)
     m_hMixer = New_hMixer
End Property ' Let hMixer

Friend Property Get Key() As String
     Key = m_Key
End Property ' Get Key

Friend Property Let Key(New_Key As String)
     m_Key = New_Key
End Property ' Let Key

Friend Property Get Maximum() As Long
     Maximum = m_Maximum
End Property ' Get Maximum

Friend Property Let Maximum(New_Maximum As Long)
     m_Maximum = New_Maximum
End Property ' Let Maximum

Friend Property Get Minimum() As Long
     Minimum = m_Minimum
End Property ' Get Minimum

Friend Property Let Minimum(New_Minimum As Long)
     m_Minimum = New_Minimum
End Property ' Let Minimum

Friend Property Get MultipleItems() As Long
     MultipleItems = m_MultipleItems
End Property ' Get MultipleItems

Friend Property Let MultipleItems(New_MultipleItems As Long)
     m_MultipleItems = New_MultipleItems
End Property ' Let MultipleItems

Friend Property Get Parent() As String
     Parent = m_Parent
End Property ' Get Parent

Friend Property Let Parent(New_Parent As String)
     m_Parent = New_Parent
End Property ' Let Parent

Friend Property Get Value() As Long
     Value = m_Value
End Property ' Get Value

Friend Property Let Value(New_Value As Long)
     m_Value = New_Value
End Property ' Let Value

Friend Function GetControlValue(ByRef lValue As Long) As Long
     Dim mcd As MIXERCONTROLDETAILS
     Dim mcdu As MIXERCONTROLDETAILS_UNSIGNED
     Dim lCnt As Long
     Dim lRtn As Long
     Dim hMem As Long
     Dim lSize As Long
     Dim lLoop As Long
     Dim lMixer As Long
     ' is multiple items?
     lCnt = Me.MultipleItems
     ' If multiple items exit.  I can not support multiple items until
     ' I have a sound card containing multiple items which make sense to me.
     ' In other words, I am unsure of implementation at the moment.  Perhaps
     ' someone can enlighten me.  I can get the multiple items in an array
     ' of structs but what is returned does not have any text labels when
     ' I use the MIXERCONTROLDETAILS_LISTTEXT struct and I do not know how
     ' to represent these items visually on the form.  Anyway......
     If lCnt > False Then Exit Function
     ' Get mixer handle
     lMixer = Me.hMixer
     ' Initialize MIXERCONTROLDETAILS struct
     With mcd
          ' Set size of struct
          .cbStruct = Len(mcd)
          ' Set the control id
          .dwControlID = Me.ControlID
           ' how many channels to return values.  I set to 1 to get uniform values
           ' for each channels of a control.  In the future will implement each channel
           ' separately and add a balance control.
          .cChannels = 1 ' Me.Channels
          ' Set size of the struct
          .cbDetails = Len(mcdu)
          ' Allocate memory to hold returned struct(s)
          hMem = GlobalAlloc(&H40, Len(mcdu))
          ' set the memory pointer
          .paDetails = GlobalLock(hMem)
     End With
     ' Call api
     lRtn = mixerGetControlDetails(lMixer, mcd, mcds.MIXER_GETCONTROLDETAILSF_VALUE)
     ' If successful
     If lRtn = MMSYSERR_NOERROR Then
          ' Get the struct from memory and get the control value
          CopyStructFromPtr mcdu, mcd.paDetails, Len(mcdu)
          ' Return value
          lValue = mcdu.dwValue
     End If
     ' Free allocated memory
     GlobalFree (hMem)
     ' Return function
     GetControlValue = lRtn
End Function ' GetLineControlDetails

Friend Function SetControlValue(ByVal lValue As Long) As Long
     Dim mcd As MIXERCONTROLDETAILS
     Dim mcdu As MIXERCONTROLDETAILS_UNSIGNED
     Dim lCnt As Long
     Dim lRtn As Long
     Dim hMem As Long
     Dim lMixer As Long
     ' is multiple items?
     lCnt = Me.MultipleItems
     ' Initialize MIXERCONTROLDETAILS struct
     ' If multiple items exit.  I can not support multiple items until
     ' I have a sound card containing multiple items which make sense to me.
     ' In other words, I am unsure of implementation at the moment.  Perhaps
     ' someone can enlighten me.  I can get the multiple items in an array
     ' of structs but what is returned does not have any text labels when
     ' I use the MIXERCONTROLDETAILS_LISTTEXT struct and I do not know how
     ' to represent these items visually on the form.  Anyway......
     If lCnt > False Then Exit Function
     ' Get mixer handle
     lMixer = Me.hMixer
     ' Initialize MIXERCONTROLDETAILS struct
     With mcd
          ' Set size of struct
          .cbStruct = Len(mcd)
          ' Set the control id
          .dwControlID = Me.ControlID
           ' how many channels to return values.  I set to 1 to set uniform values
           ' for each channel of the control.  In the future will implement each channel
           ' separately and add a balance control.
          .cChannels = 1 ' Me.Channels
          ' Set size of the struct
          .cbDetails = Len(mcdu)
          ' set the size of memory to allocate for the
          ' returned struct
          hMem = GlobalAlloc(&H40, Len(mcdu))
          ' set the memory pointer
          .paDetails = GlobalLock(hMem)
     End With
     ' Set value within MIXERCONTROLDETAILS_UNSIGNED struct
     mcdu.dwValue = lValue
     ' Walk the memory and set the control value
     CopyPtrFromStruct mcd.paDetails, mcdu, Len(mcdu)
     ' Set the value
     lRtn = mixerSetControlDetails(lMixer, mcd, MIXER_SETCONTROLDETAILSF_VALUE)
     ' Free allocated memory
     GlobalFree (hMem)
     ' Return function
     SetControlValue = lRtn
End Function ' SetLineControlDetails
