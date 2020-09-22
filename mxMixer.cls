VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mxMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'**************************************************************************************************
'  Copyright © 2003, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************

'**************************************************************************************************
' mxMixer Property Variables
'**************************************************************************************************
Private m_Destinations As Long
Private m_DeviceID As Long
Private m_DriverVersion As String
Private m_hMixer As Long
Private m_Initialized As Boolean
Private m_Key As String
Private m_ManufacturerID As Integer
Private m_ProductID As Integer
Private m_ProductName As String
Private m_MixerErr As String

'**************************************************************************************************
' mxMixer Read Properties
'**************************************************************************************************
Public Property Get Destinations() As Long
     Destinations = m_Destinations
End Property ' Get Destinations

Public Property Get DeviceID() As Long
     DeviceID = m_DeviceID
End Property ' Get DeviceID

Public Property Get DriverVersion() As String
     DriverVersion = m_DriverVersion
End Property ' Get Version

Public Property Get hMixer() As Long
     hMixer = m_hMixer
End Property ' Get hMixer

Public Property Get Initialized() As Boolean
     Initialized = m_Initialized
End Property ' Get Initialized

Public Property Get Key() As String
     Key = m_Key
End Property ' Key

Public Property Get ManufacturerID() As Integer
     ManufacturerID = m_ManufacturerID
End Property ' Get ManufacturerID

Public Property Get ProductID() As Integer
     ProductID = m_ProductID
End Property ' Get ProductID

Public Property Get ProductName() As String
     ProductName = m_ProductName
End Property ' Get ProductName

'**************************************************************************************************
' mxMixer Public Methods
'**************************************************************************************************
Public Function InitMixer(lhWnd As Long, bCallBack As Boolean, _
     Optional lMxID As Long = 0) As Long
     Dim lRtn As Long
     ' Populate mixer info
     lRtn = PopulateMixer(lMxID, bCallBack, lhWnd)
     ' if successful
     If MMSYSERR_NOERROR = 0 Then m_Initialized = True
     ' Set Return
     InitMixer = lRtn
End Function ' InitMixer

'**************************************************************************************************
' mxMixer Private Methods
'**************************************************************************************************
Private Function PopulateMixer(lMxID As Long, bCallBack As Boolean, lhWnd As Long) As Long
     Dim lRtn As Long
     Dim lMixer As Long
     Dim mc As MIXERCAPS
     ' Is class already initialized
     If Me.Initialized Then TermMixer
     ' If we want a callback
     If bCallBack Then
          ' open mixer using with callback flag and window handle
          lRtn = mixerOpen(lMixer, lMxID, lhWnd, 0&, CALLBACK_WINDOW)
     Else
          ' open mixer without callback notification
          lRtn = mixerOpen(lMixer, lMxID, 0&, 0&, 0&)
     End If
     ' If successful
     If lRtn = MMSYSERR_NOERROR Then
          ' Store mixer handle
          m_hMixer = lMixer
          ' get device capabilites
          lRtn = mixerGetDevCaps(lMixer, mc, Len(mc))
          ' If successful
          If lRtn = MMSYSERR_NOERROR Then
               ' Store device capabilities
               With mc
                    m_Destinations = .cDestinations
                    m_DeviceID = lMxID
                    m_DriverVersion = CStr(HiByte(mc.vDriverVersion) & _
                                   "." & LoByte(mc.vDriverVersion))
                    m_hMixer = lMixer
                    m_Initialized = True
                    m_Key = MXR + CStr(lMxID)
                    m_ManufacturerID = .wMid
                    m_ProductID = .wPid
                    m_ProductName = TrimNulls(mc.szPname)
               End With
          End If
     End If
     ' Return function
     PopulateMixer = lRtn
End Function ' PopulateMixer

Private Sub EmptyMixer()
     ' empty properties
     m_Destinations = 0
     m_DeviceID = 0
     m_DriverVersion = ""
     m_hMixer = 0
     m_Initialized = False
     m_Key = ""
     m_ManufacturerID = 0
     m_ProductID = 0
     m_ProductName = ""
End Sub ' EmptyMixer

Private Sub TermMixer()
     ' Close mixer
     mixerClose Me.hMixer
     ' Empty
     EmptyMixer
End Sub ' TermMixer

Private Function HiByte(ByVal wParam As Integer) As Byte
     HiByte = (wParam And &HFF00&) \ (&H100)
End Function ' HiByte

Private Function LoByte(ByVal wParam As Integer) As Byte
     LoByte = wParam And &HFF&
End Function ' LoByte

Private Function TrimNulls(sString As String) As String
     TrimNulls = Left(sString, InStr(1, sString, Chr(0)) - 1)
End Function ' TrimNulls

'**************************************************************************************************
' mxMixer Intrinsic Methods
'**************************************************************************************************
Private Sub Class_Terminate()
     ' close mixer
     If Me.hMixer > False Then TermMixer
End Sub ' Class_Terminate
