Attribute VB_Name = "mdlMixerAPI"
Option Explicit
'**************************************************************************************************
'  Copyright © 2003, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************

'**************************************************************************************************
' MMSystem API General Constants taken from MMSYSTEM.H
'**************************************************************************************************
Public Const MAXPNAMELEN& = 32
Public Const MAXERRORLENGTH = 256

'**************************************************************************************************
' MMSystem API Error Constants taken from MMSYSTEM.H
'**************************************************************************************************
Public Const MMSYSERR_NOERROR = 0                          ' no error
Public Const MMSYSERR_BASE = 0                             ' no error
Public Const MMSYSERR_ERROR = (MMSYSERR_BASE + 1)          ' unspecified error
Public Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)    ' device ID out of range
Public Const MMSYSERR_NOTENABLED = (MMSYSERR_BASE + 3)     ' driver failed enable
Public Const MMSYSERR_ALLOCATED = (MMSYSERR_BASE + 4)      ' device already allocated
Public Const MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5)    ' device handle is invalid
Public Const MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)       ' no device driver present
Public Const MMSYSERR_NOMEM = (MMSYSERR_BASE + 7)          ' memory allocation error
Public Const MMSYSERR_NOTSUPPORTED = (MMSYSERR_BASE + 8)   ' function isn't supported
Public Const MMSYSERR_BADERRNUM = (MMSYSERR_BASE + 9)      ' error value out of range
Public Const MMSYSERR_INVALFLAG = (MMSYSERR_BASE + 10)     ' invalid flag passed
Public Const MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11)    ' invalid parameter passed
Public Const MMSYSERR_HANDLEBUSY = (MMSYSERR_BASE + 12)    ' handle in use by another thread
Public Const MMSYSERR_INVALIDALIAS = (MMSYSERR_BASE + 13)  ' specified alias not found
Public Const MMSYSERR_BADDB = (MMSYSERR_BASE + 14)         ' bad registry database
Public Const MMSYSERR_KEYNOTFOUND = (MMSYSERR_BASE + 15)   ' registry key not found
Public Const MMSYSERR_READERROR = (MMSYSERR_BASE + 16)     ' registry read error
Public Const MMSYSERR_WRITEERROR = (MMSYSERR_BASE + 17)    ' registry write error
Public Const MMSYSERR_DELETEERROR = (MMSYSERR_BASE + 18)   ' registry delete error
Public Const MMSYSERR_VALNOTFOUND = (MMSYSERR_BASE + 19)   ' registry value not found
Public Const MMSYSERR_NODRIVERCB = (MMSYSERR_BASE + 20)    ' driver does not call DriverCallback
Public Const MMSYSERR_LASTERROR = (MMSYSERR_BASE + 20)     ' last error in range

'**************************************************************************************************
' MMSystem Mixer Specific API Error Constants taken from MMSYSTEM.H
'**************************************************************************************************
Public Const MIXERR_BASE = 1024
Public Const MIXERR_INVALLINE = (MIXERR_BASE + 0)
Public Const MIXERR_INVALCONTROL = (MIXERR_BASE + 1)
Public Const MIXERR_INVALVALUE = (MIXERR_BASE + 2)
Public Const MIXERR_LASTERROR = (MIXERR_BASE + 2)

'**************************************************************************************************
' MMSystem API Mixer Support Constants taken from MMSYSTEM.H
'**************************************************************************************************
Public Const MIXER_LONG_NAME_CHARS = 64
Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXER_OBJECTF_HANDLE = &H80000000
Public Const MIXER_OBJECTF_MIXER = &H0
Public Const MIXER_OBJECTF_HMIXER = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)
Public Const MIXER_OBJECTF_WAVEOUT = &H10000000
Public Const MIXER_OBJECTF_HWAVEOUT = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)
Public Const MIXER_OBJECTF_WAVEIN = &H20000000
Public Const MIXER_OBJECTF_HWAVEIN = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)
Public Const MIXER_OBJECTF_MIDIOUT = &H30000000
Public Const MIXER_OBJECTF_HMIDIOUT = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)
Public Const MIXER_OBJECTF_MIDIIN = &H40000000
Public Const MIXER_OBJECTF_HMIDIIN = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIIN)
Public Const MIXER_OBJECTF_AUX = &H50000000

' This function retrieves the number of mixer devices present in the system
Public Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long

'**************************************************************************************************
' MIXERCAPS Struct taken from MMSYSTEM.H
'**************************************************************************************************
Public Type MIXERCAPS
     wMid As Integer
     wPid As Integer
     vDriverVersion As Long
     szPname As String * MAXPNAMELEN
     fdwSupport As Long
     cDestinations As Long
End Type ' MIXERCAPS

'**************************************************************************************************
' MIXER Device API taken from MMSYSTEM.H
'**************************************************************************************************
' This function closes the specified mixer device
Public Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
' This function queries a specified mixer device to determine its capabilities
Public Declare Function mixerGetDevCaps Lib "winmm.dll" Alias "mixerGetDevCapsA" ( _
     ByVal uMxId As Long, pmxcaps As MIXERCAPS, ByVal cbmxcaps As Long) As Long
Public Declare Function mixerGetID Lib "winmm.dll" (ByVal hmxobj As Long, pumxID As Long, _
     ByVal fdwId As Long) As Long
' This function opens a specified mixer device and ensures that the device will not
' be removed until the application closes the handle
Public Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, _
     ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
' This function sends a custom mixer driver message directly to a mixer driver
Public Declare Function mixerMessage Lib "winmm.dll" (ByVal hmx As Long, _
     ByVal uMsg As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long) As Long

'**************************************************************************************************
' MIXERLINE Struct taken from MMSYSTEM.H
'**************************************************************************************************
Public Type MIXERLINE
     cbStruct As Long
     dwDestination As Long
     dwSource As Long
     dwLineID As Long
     fdwLine As Long
     dwUser As Long
     dwComponentType As Long
     cChannels As Long
     cConnections As Long
     cControls As Long
     szShortName As String * MIXER_SHORT_NAME_CHARS
     szName As String * MIXER_LONG_NAME_CHARS
     dwType As Long
     dwDeviceID As Long
     wMid  As Integer
     wPid As Integer
     vDriverVersion As Long
     szPname As String * MAXPNAMELEN
End Type ' MIXERLINE

'**************************************************************************************************
' MIXERLINE.fdwLine Constants taken from MMSYSTEM.H
'**************************************************************************************************
Public Const MIXERLINE_LINEF_ACTIVE = &H1
Public Const MIXERLINE_LINEF_DISCONNECTED = &H8000
Public Const MIXERLINE_LINEF_SOURCE = &H80000000

'**************************************************************************************************
' MIXERLINE.fdwLine Constants taken from MMSYSTEM.H
'**************************************************************************************************
Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST& = &H0&
Public Const MIXERLINE_COMPONENTTYPE_DST_UNDEFINED& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 0)
Public Const MIXERLINE_COMPONENTTYPE_DST_DIGITAL& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 1)
Public Const MIXERLINE_COMPONENTTYPE_DST_LINE& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 2)
Public Const MIXERLINE_COMPONENTTYPE_DST_MONITOR& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_DST_HEADPHONES& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 5)
Public Const MIXERLINE_COMPONENTTYPE_DST_TELEPHONE& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 6)
Public Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Public Const MIXERLINE_COMPONENTTYPE_DST_VOICEIN& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)
Public Const MIXERLINE_COMPONENTTYPE_DST_LAST& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)
Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST& = &H1000&
Public Const MIXERLINE_COMPONENTTYPE_SRC_UNDEFINED& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 0)
Public Const MIXERLINE_COMPONENTTYPE_SRC_DIGITAL& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 1)
Public Const MIXERLINE_COMPONENTTYPE_SRC_LINE& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
Public Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)
Public Const MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 6)
Public Const MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 7)
Public Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Public Const MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 9)
Public Const MIXERLINE_COMPONENTTYPE_SRC_ANALOG& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 10)
Public Const MIXERLINE_COMPONENTTYPE_SRC_LAST& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 10)

'**************************************************************************************************
' MIXERLINE.dwType Constants taken from MMSYSTEM.H
'**************************************************************************************************
Public Const MIXERLINE_TARGETTYPE_UNDEFINED = 0
Public Const MIXERLINE_TARGETTYPE_WAVEOUT = 1
Public Const MIXERLINE_TARGETTYPE_WAVEIN = 2
Public Const MIXERLINE_TARGETTYPE_MIDIOUT = 3
Public Const MIXERLINE_TARGETTYPE_MIDIIN = 4
Public Const MIXERLINE_TARGETTYPE_AUX = 5

'**************************************************************************************************
' MIXERLINEINFO API (fdwInfo Flags) taken from MMSYSTEM.H
'**************************************************************************************************
Public Const MIXER_GETLINEINFOF_DESTINATION = &H0
Public Const MIXER_GETLINEINFOF_SOURCE = &H1
Public Const MIXER_GETLINEINFOF_LINEID = &H2
Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3
Public Const MIXER_GETLINEINFOF_TARGETTYPE = &H4
Public Const MIXER_GETLINEINFOF_QUERYMASK = &HF

'**************************************************************************************************
' MIXERLINEINFO API taken from MMSYSTEM.H
'**************************************************************************************************
' This function retrieves information about a specific line of a mixer device
Public Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" ( _
     ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long

'**************************************************************************************************
' MIXERCONTROL Struct taken from MMSYSTEM.H
'**************************************************************************************************
Public Type MIXERCONTROL
     cbStruct As Long
     dwControlID As Long
     dwControlType As Long
     fdwControl As Long
     cMultipleItems As Long
     szShortName As String * MIXER_SHORT_NAME_CHARS
     szName As String * MIXER_LONG_NAME_CHARS
     lMinimum As Long
     lMaximum As Long
     reserved(9) As Long
End Type ' MIXERCONTROL

'**************************************************************************************************
' MIXERCONTROL.dwControlType Constants taken from MMSYSTEM.H
'**************************************************************************************************
' General MIXERCONTROL Type Constants
Public Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Public Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Public Const MIXERCONTROL_CT_UNITS_DECIBELS = &H40000           ' /* in 10ths */
Public Const MIXERCONTROL_CT_UNITS_PERCENT = &H50000            ' /* in 10ths */

' Custom Control Type
Public Const MIXERCONTROL_CT_CLASS_CUSTOM = &H0
Public Const MIXERCONTROL_CT_UNITS_CUSTOM = &H0
Public Const MIXERCONTROL_CONTROLTYPE_CUSTOM = (MIXERCONTROL_CT_CLASS_CUSTOM Or _
     MIXERCONTROL_CT_UNITS_CUSTOM)
     
' Fader Control Type
Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or _
     MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_BASS = (MIXERCONTROL_CONTROLTYPE_FADER + 2)
Public Const MIXERCONTROL_CONTROLTYPE_TREBLE = (MIXERCONTROL_CONTROLTYPE_FADER + 3)
Public Const MIXERCONTROL_CONTROLTYPE_EQUALIZER = (MIXERCONTROL_CONTROLTYPE_FADER + 4)

' List Control Type
Public Const MIXERCONTROL_CT_SC_LIST_SINGLE = &H0
Public Const MIXERCONTROL_CT_SC_LIST_MULTIPLE = &H1000000
Public Const MIXERCONTROL_CT_CLASS_LIST = &H70000000
Public Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or _
     MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_MIXER = (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
Public Const MIXERCONTROL_CONTROLTYPE_SINGLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or _
     MIXERCONTROL_CT_SC_LIST_SINGLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_MUX = (MIXERCONTROL_CONTROLTYPE_SINGLESELECT + 1)

' Meter Control Type
Public Const MIXERCONTROL_CT_SC_METER_POLLED = &H0
Public Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEANMETER = (MIXERCONTROL_CT_CLASS_METER Or _
     MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or _
     MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_UNSIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or _
     MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_UNSIGNED)

' Number Control Type
Public Const MIXERCONTROL_CT_CLASS_NUMBER = &H30000000
Public Const MIXERCONTROL_CONTROLTYPE_DECIBELS = (MIXERCONTROL_CT_CLASS_NUMBER Or _
     MIXERCONTROL_CT_UNITS_DECIBELS)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or _
     MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_UNSIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or _
     MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PERCENT = (MIXERCONTROL_CT_CLASS_NUMBER Or _
     MIXERCONTROL_CT_UNITS_PERCENT)

' Slider Control Type
Public Const MIXERCONTROL_CT_CLASS_SLIDER = &H40000000
Public Const MIXERCONTROL_CONTROLTYPE_SLIDER = (MIXERCONTROL_CT_CLASS_SLIDER Or _
     MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_QSOUNDPAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 2)

' Switch Control Type
Public Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0
Public Const MIXERCONTROL_CT_SC_SWITCH_BUTTON = &H1000000
Public Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or _
     MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_ONOFF = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 1)
Public Const MIXERCONTROL_CONTROLTYPE_MUTE = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
Public Const MIXERCONTROL_CONTROLTYPE_MONO = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 3)
Public Const MIXERCONTROL_CONTROLTYPE_LOUDNESS = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 4)
Public Const MIXERCONTROL_CONTROLTYPE_STEREOENH = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 5)
Public Const MIXERCONTROL_CONTROLTYPE_BUTTON = (MIXERCONTROL_CT_CLASS_SWITCH Or _
     MIXERCONTROL_CT_SC_SWITCH_BUTTON Or MIXERCONTROL_CT_UNITS_BOOLEAN)

' Time Control Types
Public Const MIXERCONTROL_CT_SC_TIME_MICROSECS = &H0
Public Const MIXERCONTROL_CT_SC_TIME_MILLISECS = &H1000000
Public Const MIXERCONTROL_CT_CLASS_TIME = &H60000000
Public Const MIXERCONTROL_CONTROLTYPE_MICROTIME = (MIXERCONTROL_CT_CLASS_TIME Or _
     MIXERCONTROL_CT_SC_TIME_MICROSECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_MILLITIME = (MIXERCONTROL_CT_CLASS_TIME Or _
     MIXERCONTROL_CT_SC_TIME_MILLISECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)

'**************************************************************************************************
' MIXERCONTROL.fdwControl Constants taken from MMSYSTEM.H
'**************************************************************************************************
Public Const MIXERCONTROL_CONTROLF_UNIFORM = &H1
Public Const MIXERCONTROL_CONTROLF_MULTIPLE = &H2
Public Const MIXERCONTROL_CONTROLF_DISABLED = &H80000000

'**************************************************************************************************
' MIXERLINECONTROLS Struct taken from MMSYSTEM.H
'**************************************************************************************************
Public Type MIXERLINECONTROLS
     cbStruct As Long
     dwLineID As Long
     dwControl As Long
     cControls As Long
     cbmxctrl As Long
     pamxctrl As Long
End Type ' MIXERLINECONTROLS

'**************************************************************************************************
' MIXERGETLINECONTROLS API (fdwControls Flags) taken from MMSYSTEM.H
'**************************************************************************************************
Public Const MIXER_GETLINECONTROLSF_ALL = &H0
Public Const MIXER_GETLINECONTROLSF_ONEBYID = &H1
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2
Public Const MIXER_GETLINECONTROLSF_QUERYMASK = &HF

'**************************************************************************************************
' MIXERGETLINECONTROLS API taken from MMSYSTEM.H
'**************************************************************************************************
' This function retrieves one or more controls associated with an audio line
Public Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" ( _
     ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long

'**************************************************************************************************
' MIXERCONTROLDETAILS Struct taken from MMSYSTEM.H
'**************************************************************************************************
Public Type MIXERCONTROLDETAILS
     cbStruct As Long
     dwControlID As Long
     cChannels As Long
     Item As Long
     cbDetails As Long
     paDetails As Long
End Type ' MIXERCONTROLDETAILS

'**************************************************************************************************
' MIXERCONTROLDETAILS.cChannels Constants (USER DEFINED)
'**************************************************************************************************
' Use this flag if control type MIXERCONTROL_CONTROLTYPE_CUSTOM
Public Const CCHANNELS_CUSTOM = &H0
' Use this flag if control type MIXERCONTROL_CONTROLF_UNIFORM
Public Const CCHANNELS_UNIFORM = &H1

'**************************************************************************************************
' MIXERCONTROLDETAILS.cMultipleItems Constants (USER DEFINED)
'**************************************************************************************************
' Use this value for all controls except for a MIXERCONTROL_CONTROLF_MULTIPLE or a
' MIXERCONTROL_CONTROLTYPE_CUSTOM control.  When using a MIXERCONTROL_CONTROLTYPE_CUSTOM control
' without the MIXERCONTROL_CONTROLTYPE_CUSTOM flag, specify zero for this flag.
Public Const CMULTIPLEITEMS_ALL = &H0

'**************************************************************************************************
' MIXERCONTROLDETAILS_BOOLEAN Struct taken from MMSYSTEM.H
'**************************************************************************************************
Public Type MIXERCONTROLDETAILS_BOOLEAN
     fValue As Long
End Type ' MIXERCONTROLDETAILS_BOOLEAN

'**************************************************************************************************
' MIXERCONTROLDETAILS_LISTTEXT Struct taken from MMSYSTEM.H
'**************************************************************************************************
Public Type MIXERCONTROLDETAILS_LISTTEXT
     dwParam1 As Long
     dwParam2 As Long
     szName As String * MIXER_LONG_NAME_CHARS
End Type ' MIXERCONTROLDETAILS_LISTTEXT

'**************************************************************************************************
' MIXERCONTROLDETAILS_SIGNED Struct taken from MMSYSTEM.H
'**************************************************************************************************
Public Type MIXERCONTROLDETAILS_SIGNED
     lValue As Long
End Type ' MIXERCONTROLDETAILS_SIGNED

'**************************************************************************************************
' MIXERCONTROLDETAILS_UNSIGNED Struct taken from MMSYSTEM.H
'**************************************************************************************************
Public Type MIXERCONTROLDETAILS_UNSIGNED
     dwValue As Long
End Type ' MIXERCONTROLDETAILS_UNSIGNED

'**************************************************************************************************
' MIXERGETCONTROLDETAILS API (fdwDetails Flags) taken from MMSYSTEM.H
'**************************************************************************************************
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0
Public Const MIXER_GETCONTROLDETAILSF_LISTTEXT = &H1
Public Const MIXER_GETCONTROLDETAILSF_QUERYMASK = &HF

'**************************************************************************************************
' MIXERGETCONTROLDETAILS API taken from MMSYSTEM.H
'**************************************************************************************************
' This function retrieves details about a single control associated with an audio line.
Public Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" ( _
     ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long

'**************************************************************************************************
' MIXERGETCONTROLDETAILS API (fdwDetails Flags) taken from MMSYSTEM.H
'**************************************************************************************************
Public Const MIXER_SETCONTROLDETAILSF_VALUE = &H0
Public Const MIXER_SETCONTROLDETAILSF_LISTTEXT = &H1
Public Const MIXER_SETCONTROLDETAILSF_QUERYMASK = &HF

'**************************************************************************************************
' MIXERSETCONTROLDETAILS API taken from MMSYSTEM.H
'**************************************************************************************************
' This function sets properties of a single control associated with an audio line
Public Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Long, _
     pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
