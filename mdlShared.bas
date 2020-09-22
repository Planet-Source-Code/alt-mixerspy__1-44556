Attribute VB_Name = "mdlMethods"
Option Explicit
'**************************************************************************************************
'  Copyright © 2003, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************

'**************************************************************************************************
' Shared constants
'**************************************************************************************************
' Collection classes
Public m_colMx As mxMixers
Public m_colLn As mxLines
Public m_colCtl As mxControls
' Constants used to construct nodes keys for storing in the treeview control
' and the mixer, line, and control collections
Public Const MXR As String = "MIXER"
Public Const DLN As String = "DESTINATIONLINE"
Public Const SLN As String = "SOURCELINE"
Public Const CTL As String = "CONTROL"
' Constants for window procs
Private Const GWL_WNDPROC = (-4)
Private Const WM_SETFOCUS = &H7

'**************************************************************************************************
' Shared enums
'**************************************************************************************************
Public Enum mgli ' mixergetlineinfo
     MIXER_GETLINEINFOF_DESTINATION = &H0
     MIXER_GETLINEINFOF_SOURCE = &H1
     MIXER_GETLINEINFOF_LINEID = &H2
     MIXER_GETLINEINFOF_COMPONENTTYPE = &H3
     MIXER_GETLINEINFOF_TARGETTYPE = &H4
End Enum ' mgli

Public Enum mglc ' mixergetlinecontrols
     MIXER_GETLINECONTROLSF_ALL = &H0
     MIXER_GETLINECONTROLSF_ONEBYID = &H1
     MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2
End Enum ' mglc

Public Enum mcds ' mixercontroldetails
     MIXER_GETCONTROLDETAILSF_VALUE = &H0
     MIXER_GETCONTROLDETAILSF_LISTTEXT = &H1
End Enum ' mcd

'**************************************************************************************************
' Win32 API
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
'Window API Declarations
Private Declare Function SetWindowLong Lib "user32" Alias _
    "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias _
    "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias _
    "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
    ByVal Msg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long

'**************************************************************************************************
' Mixer Callback Constants
'**************************************************************************************************
Public Const CALLBACK_WINDOW = &H10000
Public Const MM_MIXM_CONTROL_CHANGE = &H3D1
' Callback methods
Public lpPrevWindowProc As Long
Public lpPrevSliderProc As Long
Public lpPrevCheckProc As Long
Public lpPrevButtonProc As Long
' Tracks node selection
Public nodeSel As Node

Function EnumMixerLines(ByRef mxCol As mxMixers, ByRef lnCol As mxLines) As Long
     Dim mxMxr As mxMixer
     Dim hMixer As Long
     Dim lDstLines As Long
     Dim lLoopDst As Long
     Dim ml As MIXERLINE
     Dim lRtn As Long
     Dim mxLn As mxLine
     Dim lCnt As Long
     Dim sKey As String
     Dim lSrcLines As Long
     Dim lLoopSrc As Long
     Dim iType As Integer
     ' Loop through mixers
     For Each mxMxr In mxCol
          ' Get the mixer handle
          hMixer = mxMxr.hMixer
          ' Get the number of destination lines
          lDstLines = mxMxr.Destinations
          ' Loop and get the destination lines (lines are 0 based)
          For lLoopDst = 0 To lDstLines - 1
               ' Get info for the destination line
               lRtn = GetLineInfo(hMixer, lLoopDst, _
                    mgli.MIXER_GETLINEINFOF_DESTINATION, ml)
               ' If successful
               If lRtn = MMSYSERR_NOERROR Then
                    ' Create an instance of the line class
                    Set mxLn = New mxLine
                    ' Get the line collection count
                    lCnt = lnCol.LineCount
                    ' construct a unique key for the destination line
                    sKey = DLN + CStr(lCnt)
                    ' get the line info from the mixerline struct
                    ' and populate the line class
                    With mxLn
                         .Channels = ml.cChannels
                         .ComponentType = ml.dwComponentType
                         .Connections = ml.cConnections
                         .Controls = ml.cControls
                         .Destinations = ml.dwDestination
                         .DeviceID = ml.dwDeviceID
                         .DriverVersion = CStr(HiByte(ml.vDriverVersion) & _
                              "." & LoByte(ml.vDriverVersion))
                         .hMixer = hMixer
                         .Key = sKey
                         .Line = ml.fdwLine
                         .LineID = ml.dwLineID
                         .ManufacturerID = ml.wMid
                         .Name = TrimNulls(ml.szName)
                         .Parent = mxMxr.Key
                         .ProductID = ml.wPid
                         .ProductName = TrimNulls(ml.szPname)
                         .ShortName = TrimNulls(ml.szShortName)
                         .Source = ml.dwSource
                         .User = ml.dwUser
                         .TargetType = GetTargetType(ml.dwType)
                    End With
                    ' Add the line to the line collection
                    lnCol.LineAdd mxLn, sKey
                    ' Release line instance
                    Set mxLn = Nothing
                    ' Get the source line count (= ml.cChannels)
                    lSrcLines = ml.cConnections
                    ' For each line, we must loop through and get the source lines
                    ' that are connected to them (lines are 0 based)
                    For lLoopSrc = 0 To lSrcLines - 1
                          ' Get info for the destination line
                         lRtn = GetLineInfo(hMixer, lLoopSrc, _
                              mgli.MIXER_GETLINEINFOF_SOURCE, ml)
                         ' If successful
                         If lRtn = MMSYSERR_NOERROR Then
                              ' Create an instance of the line class
                              Set mxLn = New mxLine
                              ' Get the line collection count
                              lCnt = lnCol.LineCount
                              ' construct a unique key for the destination line
                              sKey = SLN + CStr(lCnt)
                              ' get the line info from the mixerline struct
                              ' and populate the line class
                              With mxLn
                                   .Channels = ml.cChannels
                                   .ComponentType = ml.dwComponentType
                                   .Connections = ml.cConnections
                                   .Controls = ml.cControls
                                   .Destinations = ml.dwDestination
                                   .DeviceID = ml.dwDeviceID
                                   .DriverVersion = HiByte(ml.vDriverVersion) & "." & _
                                        LoByte(ml.vDriverVersion)
                                   .hMixer = hMixer
                                   .Key = sKey
                                   .Line = ml.fdwLine
                                   .LineID = ml.dwLineID
                                   .ManufacturerID = ml.wMid
                                   .Name = TrimNulls(ml.szName)
                                   .Parent = mxMxr.Key
                                   .ProductID = ml.wPid
                                   .ProductName = TrimNulls(ml.szPname)
                                   .ShortName = TrimNulls(ml.szShortName)
                                   .Source = ml.dwSource
                                   .User = ml.dwUser
                                    iType = ml.dwType
                                   .TargetType = GetTargetType(ml.dwType)
                              End With
                              ' Add the line to the line collection
                              lnCol.LineAdd mxLn, sKey
                              ' Release line instance
                              Set mxLn = Nothing
                         End If
                    Next
               End If
          Next
     Next
     ' Return function
     EnumMixerLines = lRtn
End Function ' EnumMixerLines

Function GetLineInfo(ByVal lhMixer As Long, ByVal lLineNum As Long, _
     ByVal lFlag As mgli, ByRef ml As MIXERLINE) As Long
     ' Set size of mixerline struct
     ml.cbStruct = Len(ml)
     ' Sort out flag passed to function
     Select Case lFlag
          Case mgli.MIXER_GETLINEINFOF_COMPONENTTYPE
               ' implement
          Case mgli.MIXER_GETLINEINFOF_DESTINATION
               ' set the line number to get
               ml.dwDestination = lLineNum
          Case mgli.MIXER_GETLINEINFOF_LINEID
               ' implement
          Case mgli.MIXER_GETLINEINFOF_SOURCE
               ' Set the line number to get
               ml.dwSource = lLineNum
          Case mgli.MIXER_GETLINEINFOF_TARGETTYPE
               ' implement
     End Select
     ' call api
     GetLineInfo = mixerGetLineInfo(lhMixer, ml, lFlag)
End Function ' GetLineInfo

Function GetLineControls(ByVal lMixer As Long, ByVal lLineID As Long, ByVal lCtrls As Long, _
     ByVal lFlag As mglc, mxc() As MIXERCONTROL) As Long
     Dim mxLC As MIXERLINECONTROLS
     Dim hMem As Long
     Dim lRtn As Long
     Dim lLoop As Long
     ' Dimension mixercontrol array passed to function to the
     ' number of controls (controls 0 based)
     ReDim mxc(lCtrls - 1)
     ' Initialize MIXERLINECONTROLS structure
     With mxLC
          ' Set size of MIXERLINECONTROLS structure
          .cbStruct = Len(mxLC)
          ' Get the line id
          .dwLineID = lLineID
          ' pass control count to MIXERLINECONTROLS struct
          .cControls = lCtrls
           ' Set the size of the MIXERCONTROL struct
           .cbmxctrl = Len(mxc(0))
           ' Allocate memory to hold mixercontrol info
           hMem = GlobalAlloc(&H40, Len(mxc(0)) * .cControls)
           ' Set the pointer to allocated memory
          .pamxctrl = GlobalLock(hMem)
     End With
     ' Call api
     lRtn = mixerGetLineControls(lMixer, mxLC, lFlag)
     ' if successful
     If lRtn = MMSYSERR_NOERROR Then
          ' get array of structs stored in allocated memory
          For lLoop = 0 To lCtrls - 1
               ' Walk the memory and get each MIXERCONTROL structure
               CopyStructFromPtr mxc(lLoop), mxLC.pamxctrl + lLoop * mxLC.cbmxctrl, _
                    Len(mxc(lLoop))
          Next
     End If
     ' Free allocated memory
     GlobalFree hMem
     ' Return function
     GetLineControls = lRtn
End Function ' GetLineControls

Function EnumLineControls(ByRef lnCol As mxLines, ByRef ctlCol As mxControls) As Long
     Dim mxLn As mxLine
     Dim mxCtl As mxControl
     Dim mxc() As MIXERCONTROL
     Dim lLoop As Long
     Dim lCtlCnt As Long
     Dim hMixer As Long
     Dim lRtn As Long
     Dim lLineID As Long
     Dim sLnKey As String
     Dim lCtlCol As Long
     Dim sCtlKey As String
     Dim lVal As Long
     ' Loop through lines and get the controls
     For Each mxLn In lnCol
          ' make sure we have controls on the line so as not to run
          ' unnecessary code
          lCtlCnt = mxLn.Controls
          ' If we have controls
          If lCtlCnt > False Then
               ' Get mixer handle
               hMixer = mxLn.hMixer
               ' Get the line id
               lLineID = mxLn.LineID
               ' Get the controls
               lRtn = GetLineControls(hMixer, lLineID, lCtlCnt, _
                    MIXER_GETLINECONTROLSF_ALL, mxc)
               ' If successful
               If lRtn = MMSYSERR_NOERROR Then
                    ' Get the line key to use as parent
                    sLnKey = mxLn.Key
                    ' Loop through controls and store in class
                    For lLoop = 0 To UBound(mxc)
                         ' create control class instance
                         Set mxCtl = New mxControl
                         ' Get control collection count
                         lCtlCol = ctlCol.CtrlCount
                         ' Contstruct a key
                         sCtlKey = CTL + CStr(lCtlCol)
                         ' Set class properties
                         With mxCtl
                              .Control = mxc(lLoop).fdwControl
                              .ControlDesc = GetControlType(mxc(lLoop).dwControlType)
                              .ControlID = mxc(lLoop).dwControlID
                              .ControlName = TrimNulls(mxc(lLoop).szName)
                              .ControlShortName = TrimNulls(mxc(lLoop).szShortName)
                              .ControlType = mxc(lLoop).dwControlType
                              .hMixer = hMixer
                              .Key = sCtlKey
                              .Maximum = mxc(lLoop).lMaximum
                              .Minimum = mxc(lLoop).lMinimum
                              .MultipleItems = mxc(lLoop).cMultipleItems
                              .Parent = sLnKey
                         End With
                         ' Get value
                         mxCtl.GetControlValue lVal
                         ' Set the value for the control
                         mxCtl.Value = lVal
                         ' Add control to collection
                         ctlCol.CtrlAdd mxCtl, sCtlKey
                         ' Destroy local class instance
                         Set mxCtl = Nothing
                    Next
               End If
          End If
     Next
     ' Return function
     EnumLineControls = lRtn
End Function ' EnumLineControls

Function GetMixerError(lErr As Long) As String
     Dim sRtn As String
     Select Case lErr
          Case MMSYSERR_ERROR
               sRtn = "An unspecified error has occurred."
          Case MMSYSERR_BADDEVICEID
               sRtn = "The lID parameter of the EnumMixers function specifies an " + _
                    "device identifier out of range."
          Case MMSYSERR_NOTENABLED
               sRtn = "The device driver is not enabled."
          Case MMSYSERR_ALLOCATED
               sRtn = "The device is already allocated to the maximum number of connections."
          Case MMSYSERR_INVALHANDLE
               sRtn = "The uMxId parameter of the mixerOpen function specifies an invalid handle."
          Case MMSYSERR_NODRIVER
               sRtn = "No device driver is present."
          Case MMSYSERR_NOMEM
               sRtn = "Unable to allocate memory resources."
          Case MMSYSERR_NOTSUPPORTED
               sRtn = "Called function is not supported."
          Case MIXERR_INVALLINE
               sRtn = "The audio line reference is invalid."
          Case MMSYSERR_INVALFLAG
               sRtn = "A flag passed to the mixerOpen function is invalid."
          Case MMSYSERR_INVALPARAM
               sRtn = "One or more parameters passed to the mixerOpen function are invalid."
          Case MMSYSERR_NODRIVER
               sRtn = "No mixer device is available for the object specified by the lID parameter."
     End Select
     ' set error string
     GetMixerError = sRtn
End Function ' GetMixerError

Function GetControlType(lType As Long) As String
     Dim sRtn As String
     Select Case lType
          Case MIXERCONTROL_CONTROLTYPE_CUSTOM
               sRtn = "MIXERCONTROL_CONTROLTYPE_CUSTOM"
          Case MIXERCONTROL_CONTROLTYPE_BOOLEANMETER
               sRtn = "MIXERCONTROL_CONTROLTYPE_BOOLEANMETER"
          Case MIXERCONTROL_CONTROLTYPE_SIGNEDMETER
               sRtn = "MIXERCONTROL_CONTROLTYPE_SIGNEDMETER"
          Case MIXERCONTROL_CONTROLTYPE_PEAKMETER
               sRtn = "MIXERCONTROL_CONTROLTYPE_PEAKMETER"
          Case MIXERCONTROL_CONTROLTYPE_UNSIGNEDMETER
               sRtn = "MIXERCONTROL_CONTROLTYPE_UNSIGNEDMETER"
          Case MIXERCONTROL_CONTROLTYPE_BOOLEAN
               sRtn = "MIXERCONTROL_CONTROLTYPE_BOOLEAN"
          Case MIXERCONTROL_CONTROLTYPE_ONOFF
               sRtn = "MIXERCONTROL_CONTROLTYPE_ONOFF"
          Case MIXERCONTROL_CONTROLTYPE_MUTE
               sRtn = "MIXERCONTROL_CONTROLTYPE_MUTE"
          Case MIXERCONTROL_CONTROLTYPE_MONO
               sRtn = "MIXERCONTROL_CONTROLTYPE_MONO"
          Case MIXERCONTROL_CONTROLTYPE_LOUDNESS
               sRtn = "MIXERCONTROL_CONTROLTYPE_LOUDNESS"
          Case MIXERCONTROL_CONTROLTYPE_STEREOENH
               sRtn = "MIXERCONTROL_CONTROLTYPE_STEREOENH"
          Case MIXERCONTROL_CONTROLTYPE_BUTTON
               sRtn = "MIXERCONTROL_CONTROLTYPE_BUTTON"
          Case MIXERCONTROL_CONTROLTYPE_DECIBELS
               sRtn = "MIXERCONTROL_CONTROLTYPE_DECIBELS"
          Case MIXERCONTROL_CONTROLTYPE_SIGNED
               sRtn = "MIXERCONTROL_CONTROLTYPE_SIGNED"
          Case MIXERCONTROL_CONTROLTYPE_UNSIGNED
               sRtn = "MIXERCONTROL_CONTROLTYPE_UNSIGNED"
          Case MIXERCONTROL_CONTROLTYPE_PERCENT
               sRtn = "MIXERCONTROL_CONTROLTYPE_PERCENT"
          Case MIXERCONTROL_CONTROLTYPE_SLIDER
               sRtn = "MIXERCONTROL_CONTROLTYPE_SLIDER"
          Case MIXERCONTROL_CONTROLTYPE_PAN
               sRtn = "MIXERCONTROL_CONTROLTYPE_PAN"
          Case MIXERCONTROL_CONTROLTYPE_QSOUNDPAN
               sRtn = "MIXERCONTROL_CONTROLTYPE_QSOUNDPAN"
          Case MIXERCONTROL_CONTROLTYPE_FADER
               sRtn = "MIXERCONTROL_CONTROLTYPE_FADER"
          Case MIXERCONTROL_CONTROLTYPE_VOLUME
               sRtn = "MIXERCONTROL_CONTROLTYPE_VOLUME"
          Case MIXERCONTROL_CONTROLTYPE_BASS
               sRtn = "MIXERCONTROL_CONTROLTYPE_BASS"
          Case MIXERCONTROL_CONTROLTYPE_TREBLE
               sRtn = "MIXERCONTROL_CONTROLTYPE_TREBLE"
          Case MIXERCONTROL_CONTROLTYPE_EQUALIZER
               sRtn = "MIXERCONTROL_CONTROLTYPE_EQUALIZER"
          Case MIXERCONTROL_CONTROLTYPE_SINGLESELECT
               sRtn = "MIXERCONTROL_CONTROLTYPE_SINGLESELECT"
          Case MIXERCONTROL_CONTROLTYPE_MUX
               sRtn = "MIXERCONTROL_CONTROLTYPE_MUX"
          Case MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT
               sRtn = "MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT"
          Case MIXERCONTROL_CONTROLTYPE_MIXER
               sRtn = "MIXERCONTROL_CONTROLTYPE_MIXER"
          Case MIXERCONTROL_CONTROLTYPE_MICROTIME
               sRtn = "MIXERCONTROL_CONTROLTYPE_MICROTIME"
          Case MIXERCONTROL_CONTROLTYPE_MILLITIME
               sRtn = "MIXERCONTROL_CONTROLTYPE_MILLITIME"
          Case Else
               sRtn = "UNKNOWN"
     End Select
     GetControlType = sRtn
End Function ' GetControlType

Function GetLineComponent(lType As Long) As String
     Dim sRtn As String
     Select Case lType
          Case MIXERLINE_COMPONENTTYPE_DST_UNDEFINED
               sRtn = "DST_UNDEFINED"
          Case MIXERLINE_COMPONENTTYPE_DST_DIGITAL
               sRtn = "DST_DIGITAL"
          Case MIXERLINE_COMPONENTTYPE_DST_LINE
               sRtn = "DST_LINE"
          Case MIXERLINE_COMPONENTTYPE_DST_MONITOR
               sRtn = "DST_MONITOR"
          Case MIXERLINE_COMPONENTTYPE_DST_SPEAKERS
               sRtn = "DST_SPEAKERS"
          Case MIXERLINE_COMPONENTTYPE_DST_SPEAKERS
               sRtn = "DST_SPEAKERS"
          Case MIXERLINE_COMPONENTTYPE_DST_TELEPHONE
               sRtn = "DST_TELEPHONE"
          Case MIXERLINE_COMPONENTTYPE_DST_WAVEIN
               sRtn = "DST_WAVEIN"
          Case MIXERLINE_COMPONENTTYPE_DST_VOICEIN
               sRtn = "DST_VOICEIN"
          Case MIXERLINE_COMPONENTTYPE_SRC_UNDEFINED
               sRtn = "SRC_UNDEFINED"
          Case MIXERLINE_COMPONENTTYPE_SRC_DIGITAL
               sRtn = "SRC_DIGITAL"
          Case MIXERLINE_COMPONENTTYPE_SRC_LINE
               sRtn = "SRC_LINE"
          Case MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE
               sRtn = "SRC_MICROPHONE"
          Case MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER
               sRtn = "SRC_SYNTHESIZER"
          Case MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC
               sRtn = "SRC_COMPACTDISC"
          Case MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE
               sRtn = "SRC_TELEPHONE"
          Case MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER
               sRtn = "SRC_PCSPEAKER"
          Case MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT
               sRtn = "SRC_WAVEOUT"
          Case MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY
               sRtn = "SRC_AUXILIARY"
          Case MIXERLINE_COMPONENTTYPE_SRC_ANALOG
               sRtn = "SRC_ANALOG"
          Case Else
               sRtn = "Source/Destinaton Unknown"
     End Select
     GetLineComponent = sRtn
End Function ' GetLineComponent

Public Function GetMixerVersion(lVer As Integer) As Integer
     Dim byHi As Byte
     Dim byLo As Byte
     byHi = HiByte(lVer)
     byLo = LoByte(lVer)
End Function ' GetMixerVersion

Public Function GetTargetType(ByVal iType As Integer) As String
     Dim sRtn As String
     Select Case iType
          Case MIXERLINE_TARGETTYPE_UNDEFINED
               sRtn = "MIXERLINE_TARGETTYPE_UNDEFINED"
          Case MIXERLINE_TARGETTYPE_WAVEOUT
               sRtn = "MIXERLINE_TARGETTYPE_WAVEOUT"
          Case MIXERLINE_TARGETTYPE_WAVEIN
               sRtn = "MIXERLINE_TARGETTYPE_WAVEIN"
          Case MIXERLINE_TARGETTYPE_MIDIOUT
               sRtn = "MIXERLINE_TARGETTYPE_MIDIOUT"
          Case MIXERLINE_TARGETTYPE_MIDIIN
               sRtn = "MIXERLINE_TARGETTYPE_MIDIIN"
          Case MIXERLINE_TARGETTYPE_AUX
               sRtn = "MIXERLINE_TARGETTYPE_AUX"
     End Select
     GetTargetType = sRtn
End Function ' GetTargetType

Public Function HiByte(ByVal wParam As Integer) As Byte
     HiByte = (wParam And &HFF00&) \ (&H100)
End Function ' HiByte

Public Function LoByte(ByVal wParam As Integer) As Byte
     LoByte = wParam And &HFF&
End Function ' LoByte

Public Function TrimNulls(sString As String) As String
     TrimNulls = Left(sString, InStr(1, sString, Chr(0)) - 1)
End Function ' TrimNulls

Public Sub HookObject(ByRef ctlObj As Object)
     Dim hWnd As Long
     Dim sName As String
     ' get object window handle
     hWnd = ctlObj.hWnd
     ' Get the name of the object
     sName = TypeName(ctlObj)
     ' Get control's window proc
     Select Case sName
          Case "CommandButton"
               lpPrevButtonProc = SetWindowLong(hWnd, GWL_WNDPROC, _
                    AddressOf ButtonProc)
          Case "CheckBox"
               lpPrevCheckProc = SetWindowLong(hWnd, GWL_WNDPROC, _
                    AddressOf CheckProc)
          Case "frmMixerSpy"
               lpPrevWindowProc = SetWindowLong(hWnd, GWL_WNDPROC, _
                    AddressOf WindowProc)
          Case "Slider"
               lpPrevSliderProc = SetWindowLong(hWnd, GWL_WNDPROC, _
                    AddressOf SliderProc)
     End Select
End Sub ' HookObject

Public Function ButtonProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, _
     ByVal lparam As Long) As Long
    'The procedure that gets all windows messages for the subclassed button
    On Error Resume Next
    Select Case uMsg&
        'The button is going to get the focus
        Case WM_SETFOCUS
        'Exit the procedure -> The message doesn´t reach the button
        Exit Function
    End Select
    'Call the standard Button Procedure
    ButtonProc = CallWindowProc(lpPrevButtonProc, hWnd, uMsg, wParam, lparam)
End Function ' ButtonProc

Public Function SliderProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, _
     ByVal lparam As Long) As Long
    'The procedure that gets all windows messages for the subclassed button
    On Error Resume Next
    Select Case uMsg&
        'The slider is going to get the focus
        Case WM_SETFOCUS
        'Exit the procedure -> The message doesn´t reach the button
        Exit Function
    End Select
    'Call the standard Button Procedure
    SliderProc = CallWindowProc(lpPrevSliderProc, hWnd, uMsg, wParam, lparam)
End Function ' SliderProc

Public Function CheckProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, _
     ByVal lparam As Long) As Long
    'The procedure that gets all windows messages for the subclassed button
    On Error Resume Next
    Select Case uMsg
        'The button is going to get the focus
        Case WM_SETFOCUS
        'Exit the procedure -> The message doesn´t reach the button
        Exit Function
    End Select
    'Call the standard Button Procedure
    CheckProc = CallWindowProc(lpPrevCheckProc, hWnd, uMsg, wParam, lparam)
End Function ' CheckProc

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, _
     ByVal lparam As Long) As Long
     Dim sItemKey As String
     Dim sSelKey As String
     Dim lVal As Long
     Dim lRtn As Long
     Dim mxCtl As mxControl
     ' Is it the message we want
     Select Case uMsg
          Case MM_MIXM_CONTROL_CHANGE
               ' construct the key to retrieve control from collection
               sItemKey = CTL + CStr(lparam - 1)
               ' Get key of selected node
               sSelKey = nodeSel.Key
               ' If the same as selected
               If sSelKey = sItemKey Then
                    ' get control from collection
                    Set mxCtl = m_colCtl(sItemKey)
                    ' Get control's new value
                    lRtn = mxCtl.GetControlValue(lVal)
                    ' Update controls
                    With mxCtl
                         If .Maximum = 1 And .Minimum = 0 Then
                              frmMixerSpy.cmdOnOff = lVal
                         Else
                              frmMixerSpy.slFader.Value = lVal
                              frmMixerSpy.txtValue = GetPctStr(frmMixerSpy.slFader.Value, _
                                   frmMixerSpy.slFader.Max)
                              frmMixerSpy.lvwItems.ListItems("Value").SubItems(1) = CStr(lVal)
                         End If
                    End With
               End If
     End Select
     WindowProc = CallWindowProc(lpPrevWindowProc, hWnd, uMsg, wParam, lparam)
End Function ' CheckProc

Public Sub UnHookObject(ctlObj As Object)
     Dim hWnd As Long
     Dim sCtlName As String
     Dim sName As String
     ' get object window handle
     hWnd = ctlObj.hWnd
     ' Get the name of the object
     sName = ctlObj.Name
     ' Get control's window proc
     Select Case sName
          Case "CommandButton"
               Call SetWindowLong(hWnd, GWL_WNDPROC, lpPrevButtonProc)
          Case "CheckBox"
               Call SetWindowLong(hWnd, GWL_WNDPROC, lpPrevCheckProc)
          Case "frmMixerSpy"
               Call SetWindowLong(hWnd, GWL_WNDPROC, lpPrevWindowProc)
          Case "Slider"
              Call SetWindowLong(hWnd, GWL_WNDPROC, lpPrevSliderProc)
          End Select
End Sub ' UnHookObject

Public Function GetPctStr(lPart As Long, lWhole As Long) As String
     Dim lRtn As Long
     ' don't mess with zeros
     If lPart = False Or lWhole = False Then
          GetPctStr = "0" + Chr(37)
          Exit Function
     End If
     ' do the division
     lRtn = Round(lPart / lWhole * 100)
     ' Return
     GetPctStr = CStr(lRtn) + Chr(37)
End Function ' GetPctStr

