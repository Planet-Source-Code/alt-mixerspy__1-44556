VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mxLine"
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
' mxLine Property Variables
'**************************************************************************************************
Private m_Channels As Long
Private m_ComponentType As Long
Private m_Connections As Long
Private m_Controls As Long
Private m_Destinations As Long
Private m_DeviceID As Long
Private m_DriverVersion As String
Private m_hMixer As Long
Private m_Key As String
Private m_Line As Long
Private m_LineID As Long
Private m_LineType As String
Private m_ManufacturerID As Long
Private m_MixerErr As String
Private m_Name As String
Private m_Parent As String
Private m_ProductID As Long
Private m_ProductName As String
Private m_ShortName As String
Private m_Source As Long
Private m_TargetType As String
Private m_User As Long

'**************************************************************************************************
' mxLine Property Pairs
'**************************************************************************************************
Friend Property Get Channels() As Long
     Channels = m_Channels
End Property ' Get Channels

Friend Property Let Channels(New_Channels As Long)
     m_Channels = New_Channels
End Property ' Let Channels

Friend Property Get ComponentType() As Long
     ComponentType = m_ComponentType
End Property ' Get ComponentType

Friend Property Let ComponentType(New_ComponentType As Long)
     m_ComponentType = New_ComponentType
End Property ' Let ComponentType

Friend Property Get Connections() As Long
     Connections = m_Connections
End Property ' Get Connections

Friend Property Let Connections(New_Connections As Long)
     m_Connections = Connections
End Property ' Let Connections

Friend Property Get Controls() As Long
     Controls = m_Controls
End Property ' Get Controls

Friend Property Let Controls(New_Controls As Long)
     m_Controls = New_Controls
End Property ' Let Controls

Friend Property Get Destinations() As Long
     Destinations = m_Destinations
End Property ' Get Destination

Friend Property Let Destinations(New_Destinations As Long)
     m_Destinations = New_Destinations
End Property ' Let Destinations

Friend Property Get DeviceID() As Long
     DeviceID = m_DeviceID
End Property ' Get DeviceID

Friend Property Let DeviceID(New_DeviceID As Long)
     m_DeviceID = New_DeviceID
End Property ' Let DeviceID

Friend Property Get DriverVersion() As String
     DriverVersion = m_DriverVersion
End Property ' Get DriverVersion

Friend Property Let DriverVersion(New_DriverVersion As String)
     m_DriverVersion = New_DriverVersion
End Property ' Let DriverVersion

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

Friend Property Get Line() As Long
     Line = m_Line
End Property ' Get Line

Friend Property Let Line(New_Line As Long)
     m_Line = New_Line
End Property ' Let Line

Friend Property Get LineID() As Long
     LineID = m_LineID
End Property ' Get LineID

Friend Property Let LineID(New_LineID As Long)
     m_LineID = New_LineID
End Property ' Let LineId

Friend Property Get ManufacturerID() As Long
     ManufacturerID = m_ManufacturerID
End Property ' Get ManufacturerID

Friend Property Let ManufacturerID(New_ManufacturerID As Long)
     m_ManufacturerID = New_ManufacturerID
End Property ' Let ManufacturerID

Friend Property Get Name() As String
     Name = m_Name
End Property ' Get Name

Friend Property Let Name(New_Name As String)
     m_Name = New_Name
End Property ' Let Name

Friend Property Get Parent() As String
     Parent = m_Parent
End Property ' Get Parent

Friend Property Let Parent(New_Parent As String)
     m_Parent = New_Parent
End Property ' Let Parent

Friend Property Get ProductID() As Long
     ProductID = m_ProductID
End Property ' Get ProductID

Friend Property Let ProductID(New_ProductID As Long)
     m_ProductID = New_ProductID
End Property ' Let ProductID

Friend Property Get ProductName() As String
     ProductName = m_ProductName
End Property ' Get ProductName

Friend Property Let ProductName(New_ProductName As String)
     m_ProductName = New_ProductName
End Property ' Let ProductName

Friend Property Get ShortName() As String
     ShortName = m_ShortName
End Property ' Get ShortName

Friend Property Let ShortName(New_ShortName As String)
     m_ShortName = New_ShortName
End Property ' Let ShortName

Friend Property Get Source() As Long
     Source = m_Source
End Property ' Get Source

Friend Property Let Source(New_Source As Long)
     m_Source = New_Source
End Property ' Let Source

Friend Property Get TargetType() As String
     TargetType = m_TargetType
End Property ' Get TargetType

Friend Property Let TargetType(New_TargetType As String)
     m_TargetType = New_TargetType
End Property ' Let TargetType

Friend Property Get User() As Long
     User = m_User
End Property ' Get User

Friend Property Let User(New_User As Long)
     m_User = New_User
End Property ' Let User

