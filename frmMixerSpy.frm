VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMixerSpy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mixer Spy"
   ClientHeight    =   4335
   ClientLeft      =   3075
   ClientTop       =   3345
   ClientWidth     =   9390
   Icon            =   "frmMixerSpy.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ilMixers 
      Left            =   2820
      Top             =   3210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMixerSpy.frx":0442
            Key             =   "mixer"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMixerSpy.frx":214C
            Key             =   "destination"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMixerSpy.frx":2654
            Key             =   "source"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMixerSpy.frx":2B5B
            Key             =   "control"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMixerSpy.frx":2E75
            Key             =   "control1"
         EndProperty
      EndProperty
   End
   Begin VB.Frame frameInfo 
      Height          =   3360
      Left            =   3540
      TabIndex        =   3
      Top             =   60
      Width           =   5775
      Begin MSComctlLib.ListView lvwItems 
         Height          =   3045
         Left            =   105
         TabIndex        =   4
         Top             =   225
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   5371
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "List"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   6174
         EndProperty
      End
   End
   Begin VB.Frame frameView 
      Caption         =   "Mixers -> Lines -> Controls"
      Height          =   4185
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   3495
      Begin VB.CommandButton cmdCopy 
         Height          =   270
         Left            =   105
         Picture         =   "frmMixerSpy.frx":2F16
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3825
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin MSComctlLib.TreeView tvwMixers 
         Height          =   3510
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   6191
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   441
         Style           =   7
         ImageList       =   "ilMixers"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCopy 
         Caption         =   "Copy Control Types To Clipboard"
         Height          =   210
         Left            =   480
         TabIndex        =   22
         Top             =   3855
         Width           =   2445
      End
   End
   Begin VB.Frame frameCtrls 
      Height          =   960
      Left            =   3540
      TabIndex        =   2
      Top             =   3285
      Width           =   4320
      Begin VB.Frame frameCtls 
         BorderStyle     =   0  'None
         Height          =   780
         Index           =   1
         Left            =   45
         TabIndex        =   7
         Top             =   105
         Width           =   4245
         Begin VB.TextBox txtValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   285
            Left            =   2340
            TabIndex        =   16
            Text            =   "0"
            Top             =   360
            Width           =   705
         End
         Begin MSComctlLib.Slider slFader 
            Height          =   420
            Left            =   45
            TabIndex        =   17
            Top             =   330
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   741
            _Version        =   393216
            Max             =   65535
            TickFrequency   =   6535
         End
         Begin VB.Label lblControlType 
            AutoSize        =   -1  'True
            Caption         =   "Control Type:  Volume"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   19
            Top             =   120
            Width           =   1590
         End
         Begin VB.Label lblValue 
            Caption         =   "Value:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1845
            TabIndex        =   18
            Top             =   405
            Width           =   495
         End
      End
      Begin VB.Frame frameCtls 
         BorderStyle     =   0  'None
         Height          =   780
         Index           =   2
         Left            =   45
         TabIndex        =   9
         Top             =   105
         Width           =   4230
         Begin VB.CheckBox cmdOnOff 
            Caption         =   "Off"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   105
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   405
            Width           =   510
         End
         Begin VB.Label lblBoolean 
            AutoSize        =   -1  'True
            Caption         =   "Control Type:  On/Off"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   15
            Top             =   120
            Width           =   1590
         End
         Begin VB.Label lblDescription 
            AutoSize        =   -1  'True
            Caption         =   "On/Off Description"
            Height          =   195
            Left            =   675
            TabIndex        =   14
            Top             =   450
            Width           =   1335
         End
      End
      Begin VB.Frame frameCtls 
         BorderStyle     =   0  'None
         Height          =   780
         Index           =   0
         Left            =   45
         TabIndex        =   8
         Top             =   105
         Width           =   4245
      End
      Begin VB.Frame frameCtls 
         BorderStyle     =   0  'None
         Height          =   780
         Index           =   3
         Left            =   45
         TabIndex        =   10
         Top             =   105
         Width           =   4230
         Begin VB.CheckBox chkMute 
            Caption         =   "Mute"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   11
            Top             =   390
            Width           =   660
         End
         Begin VB.Label lblMute 
            AutoSize        =   -1  'True
            Caption         =   "Control Type:  Mute"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   12
            Top             =   120
            Width           =   1440
         End
      End
   End
   Begin VB.Frame frameClose 
      Height          =   945
      Left            =   7830
      TabIndex        =   5
      Top             =   3300
      Width           =   1485
      Begin VB.CheckBox chkHook 
         Caption         =   "Set Hooks"
         Height          =   345
         Left            =   120
         TabIndex        =   20
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   135
         TabIndex        =   6
         Top             =   555
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmMixerSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**************************************************************************************************
'  Copyright © 2002, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************
' This code is the culmination of an effort to get as much information about
' what my sound card supports and what it does not.  Portions are based on the
' MyMixer project by Mr. Bobo?  Portions are based on a mixer project by an
' an anonymous author.  And portions are based on VolOcx by Gordon Lindsey of
' RJSoft.  Future versions will support the peakmeter and other controls not
' implemented in this app when I can figure out and purchase a sound card that
' implements most or all of the controls listed in the MSDN documentation.  I
' would appreciate very much feedback from those who have soundcards that support
' most or all of the controls.  You can use the copy to clipboard button to
' copy a list of the controls supported by your mixer and paste it into an
' email to the address above. Once I have whatever sound card supports the
' most controls I can update this app.
'**************************************************************************************************
'  Form/Control Intrinsic Subs/Methods
'**************************************************************************************************
Private Sub chkHook_Click()
     ' These hooks are set up to intercept focus to the slider and checkbox
     ' controls.  Also, to intercept mixer change messages so that the
     ' selected control can be updated if the value of the control is
     ' changed by another application.  Never hit the end button or set
     ' a breakpoint when hooking is enabled or it will lock up the IDE.
     ' If you want to step through the code simply uncheck chkHook.
     On Error GoTo errHandler
     If chkHook.Value = False Then
          ' Set label
          chkHook.Caption = "Set Hooks"
          ' subclass and interecept focus messages for these controls
          Call UnHookObject(frmMixerSpy)
          Call UnHookObject(cmdCopy)
          Call UnHookObject(cmdOnOff)
          Call UnHookObject(chkMute)
          Call UnHookObject(slFader)
     Else
          ' Set label
          chkHook.Caption = "Unhook"
          ' return control to these the standard window
          ' procs for these controls
          Call HookObject(frmMixerSpy)
          Call HookObject(cmdCopy)
          Call HookObject(cmdOnOff)
          Call HookObject(chkMute)
          Call HookObject(slFader)
     End If
     Exit Sub
errHandler:
     MsgBox "Error:  " & Err.Number + ".  " + Err.Description, vbOKOnly + _
          vbApplicationModal + vbExclamation, "MixerSpy Error"
End Sub ' chkHook

Private Sub cmdClose_Click()
     Unload Me
End Sub ' cmdClose_Click

Private Sub cmdCopy_Click()
     Dim mx As mxMixer
     Dim CTL As mxControl
     Dim lMixer As Long
     Dim lCtr As Long
     Dim sString As String
     For Each mx In m_colMx
          ' Get mixer handle
          lMixer = mx.hMixer
          ' Construct string
          sString = sString + "Mixer Product Name:  " + mx.ProductName + _
               Chr(13) + Chr(13)
          lCtr = False
          ' Loop through controls
          For Each CTL In m_colCtl
               lCtr = lCtr + 1
               ' if child of mixer
               If m_colCtl(lCtr).hMixer = lMixer Then
               ' Loop through controls and get control types
                    sString = sString + "Control Type:  " + _
                         GetControlType(CTL.ControlType) + Chr(13)
               End If
          Next
          sString = sString + Chr(13)
     Next
     Debug.Print sString
     Clipboard.SetText sString
     ' release objects
     Set CTL = Nothing
     Set mx = Nothing
End Sub ' cmdCopy

Private Sub cmdOnOff_Click()
     On Error GoTo errHandler
     ' Update button caption
     If cmdOnOff.Value = 0 Then
          cmdOnOff.Caption = "Off"
     Else
          cmdOnOff.Caption = "On"
     End If
     ' Set value
     SetValue cmdOnOff.Value
     ' Update the value in the listview
     lvwItems.ListItems("Value").SubItems(1) = CStr(cmdOnOff.Value)
     Exit Sub
errHandler:
     MsgBox "Error:  " & Err.Number + ".  " + Err.Description, vbOKOnly + _
          vbApplicationModal + vbExclamation, "MixerSpy Error"
End Sub ' cmdOnOff_Click

Private Sub Form_Load()
     Dim lRtn As Long
     Dim sErr As String
     On Error GoTo errHandler
     ' Load the mixers
     lRtn = LoadMixers
     ' Process return
     If lRtn = MMSYSERR_NOERROR Then
          ' Load the mixer lines
          lRtn = LoadMixerLines
          ' Process return
          If lRtn <> MMSYSERR_NOERROR Then GoTo errHandler
          ' load controls
          lRtn = LoadLineControls
          ' Process return
          If lRtn <> MMSYSERR_NOERROR Then GoTo errHandler
          ' Make sure child nodes are visible
          tvwMixers.Nodes(1).Expanded = True
          tvwMixers.Nodes(1).EnsureVisible
          ' Simlulate click to get root node info
          Call tvwMixers_NodeClick(tvwMixers.Nodes(1))
     Else
          GoTo errHandler
     End If
     Exit Sub
errHandler:
     If Err.Number = False Then
          ' Get error
          sErr = GetMixerError(lRtn)
          ' display error
          MsgBox sErr, vbOKOnly, "Mixer Error"
     Else
          MsgBox "Error:  " & Err.Number + ".  " + Err.Description, vbOKOnly + _
               vbApplicationModal + vbExclamation, "MixerSpy Error"
     End If
End Sub ' Form_Load

Private Sub Form_Unload(Cancel As Integer)
     Dim lLoop As Long
     Dim mx As mxMixer
     On Error GoTo errHandler
     ' Unhook controls and window
     chkHook = False
     ' Clear controls
     m_colCtl.CtrlsClear
     ' Release controls collection object
     Set m_colCtl = Nothing
     ' Clear lines
     m_colLn.LinesClear
     ' Release lines collection object
     Set m_colLn = Nothing
     ' Clear mixers
     m_colMx.MixersClear
     ' Release mixer object
     Set m_colMx = Nothing
     Exit Sub
errHandler:
     MsgBox "Error:  " & Err.Number + ".  " + Err.Description, vbOKOnly + _
          vbApplicationModal + vbExclamation, "MixerSpy Error"
End Sub ' Form_Unload

Private Sub lvwItems_BeforeLabelEdit(Cancel As Integer)
     ' Prevent edits to item labels
     Cancel = True
End Sub ' lvwItems_BeforeLabelEdit

Private Sub slFader_Scroll()
     On Error GoTo errHandler
     Dim itmX As ListItem
     ' Set the new value in the text box
     txtValue = GetPctStr(slFader.Value, slFader.Max)
     ' Update the value in the listview
     lvwItems.ListItems("Value").SubItems(1) = CStr(slFader.Value)
     ' Set the control value
     SetValue slFader.Value
     Exit Sub
errHandler:
     MsgBox "Error:  " & Err.Number + ".  " + Err.Description, vbOKOnly + _
          vbApplicationModal + vbExclamation, "MixerSpy Error"
End Sub ' slFader_Scroll

Private Sub tvwMixers_BeforeLabelEdit(Cancel As Integer)
     ' Prevent edits to node label
     Cancel = True
End Sub ' tvwMixers_BeforeLabelEdit

Private Sub tvwMixers_GotFocus()
     ' Return focus to selected node
     tvwMixers.Nodes(nodeSel.Index).Selected = True
End Sub ' tvwMixers_GotFocus

Private Sub tvwMixers_NodeClick(ByVal Node As MSComctlLib.Node)
     Dim sKey As String
     Dim lRtn As Long
     On Error GoTo errHandler
     ' Set public node variable to selected node
     Set nodeSel = Node
     ' Get the key assigned to the node
     sKey = Node.Key
     ' node is a mixer
     If Left(sKey, Len(MXR)) = MXR Then
          ' load mixer info
          LoadMixerCaps sKey
     ' node is a destination or source line
     ElseIf Left(sKey, Len(DLN)) = DLN Or Left(sKey, Len(SLN)) = SLN Then
          ' load line info
          LoadLineInfo sKey
     ' node is a line control
     ElseIf Left(sKey, Len(CTL)) = CTL Then
          ' load control info
          LoadControlInfo sKey
     End If
     Exit Sub
errHandler:
     MsgBox "Error:  " & Err.Number + ".  " + Err.Description, vbOKOnly + _
          vbApplicationModal + vbExclamation, "MixerSpy Error"
End Sub ' tvwMixers_NodeClick

'******************************************************************************
' LoadMixers - enums mixers on the system, adds them to the mixers collection,
' and loads them into the treeview as root nodes.
'******************************************************************************
Private Function LoadMixers() As Long
     Dim lMixerCnt As Long
     Dim lLoop As Long
     Dim mx As mxMixer
     Dim lRtn As Long
     Dim bRtn As Boolean
     Dim sKey As String
     Dim sName As String
     Dim xNode As Node
     On Error GoTo errHandler
     ' Clear all nodes in treeview
     tvwMixers.Nodes.Clear
     ' Do we have any mixers
     lMixerCnt = mixerGetNumDevs
     ' If we have mixers
     If lMixerCnt Then
          ' create our mixers collection
          Set m_colMx = New mxMixers
          ' Loop through lMixerCnt (Mixers are 0 based)
          For lLoop = 0 To lMixerCnt - 1
               ' Set mixer object
               Set mx = New mxMixer
               ' Initialize mixer
               lRtn = mx.InitMixer(Me.hWnd, True, lLoop)
               ' If successful
               If lRtn = MMSYSERR_NOERROR Then
                    ' Get the key from mixer class
                    sKey = mx.Key
                    ' Get the product name
                    sName = mx.ProductName
                    ' Add mixer to collection
                    m_colMx.MixerAdd mx, sKey
                    ' Add mixer to treeview
                    Set xNode = tvwMixers.Nodes.Add(, , sKey, sName, "mixer")
                    ' Set Node attributes
                    With xNode
                         .Bold = True
                         .ForeColor = &H0
                         .EnsureVisible
                    End With
                    ' release node object
                    Set xNode = Nothing
               Else
                    ' Release temporary mixer object
                    Set mx = Nothing
                    ' process error
                    GoTo errHandler
               End If
          Next
     Else
          ' failed so release mixers collection object
          Set m_colMx = Nothing
          ' Return unspecified error
          lRtn = MMSYSERR_ERROR
     End If
     ' Return
     LoadMixers = lRtn
     Exit Function
errHandler:
     MsgBox "Error:  " & Err.Number + ".  " + Err.Description, vbOKOnly + _
          vbApplicationModal + vbExclamation, "MixerSpy Error"
End Function ' GetMixers

'******************************************************************************
' LoadMixerLines - enums line objects, adds them to the lines collection,
' and loads them into the treeview as children of the mixer they are
' associated with.
'******************************************************************************
Private Function LoadMixerLines() As Long
     Dim mxLn As mxLine
     Dim lRtn As Long
     Dim sKey As String
     Dim sName As String
     Dim sParent As String
     Dim xNode As Node
     On Error GoTo errHandler
     ' create our lines collection
     Set m_colLn = New mxLines
     ' Enumerate the mixer lines
     lRtn = EnumMixerLines(m_colMx, m_colLn)
     ' If successful
     If lRtn = MMSYSERR_NOERROR Then
          ' Loop through lines
          For Each mxLn In m_colLn
               ' Get the key to identify node in treeview
               With mxLn
                    sKey = .Key
                    ' Get the line name
                    sName = .Name
                    ' Get node parent
                    sParent = .Parent
               End With
               ' Is it a destination line or a source line?
               If Left(sKey, Len(DLN)) = DLN Then
               ' add to treeview
                    Set xNode = tvwMixers.Nodes.Add(sParent, tvwChild, sKey, sName, "destination")
                    ' Set node to denote a destination line
                    xNode.Bold = True
                    ' Set color to distinguish from a source line
                    xNode.ForeColor = &H800000
               Else
                    ' add to treeview
                    Set xNode = tvwMixers.Nodes.Add(sParent, tvwChild, sKey, sName, "source")
                    ' Set node color to distinguish from destination line
                    xNode.ForeColor = &HFF0000
                    xNode.EnsureVisible
               End If
               ' good practice to release objects
               Set xNode = Nothing
          Next
     Else
          ' failed so release lines collection object
          Set m_colLn = Nothing
     End If
     ' Return function
     LoadMixerLines = lRtn
     Exit Function
errHandler:
     MsgBox "Error:  " & Err.Number + ".  " + Err.Description, vbOKOnly + _
          vbApplicationModal + vbExclamation, "MixerSpy Error"
End Function ' LoadMixerLines

'******************************************************************************
' LoadLineControls - enums line control objects, adds them to the controls
' collection, and loads them into the treeview as children of the line they
' are associated with.
'******************************************************************************
Private Function LoadLineControls() As Long
     Dim mxCtrl As mxControl
     Dim lRtn As Long
     Dim sKey As String
     Dim sName As String
     Dim sParent As String
     Dim xNode As Node
     On Error GoTo errHandler
     ' create our controls collection
     Set m_colCtl = New mxControls
     ' Enumerate the mixer lines
     lRtn = EnumLineControls(m_colLn, m_colCtl)
     ' If successful
     If lRtn = MMSYSERR_NOERROR Then
          ' Loop through controls
          For Each mxCtrl In m_colCtl
               With mxCtrl
                    ' Get the key to identify node in treeview
                    sKey = .Key
                    ' Get the control name
                    sName = .ControlName
                    ' Get node parent
                    sParent = .Parent
               End With
               ' add to treeview
               Set xNode = tvwMixers.Nodes.Add(sParent, tvwChild, sKey, sName, "control")
               ' Set node color to distinguish from destination line
               xNode.ForeColor = &HC0&
               ' good practice to release objects
               Set xNode = Nothing
          Next
     Else
          ' failed so release controls collection object
          Set m_colCtl = Nothing
     End If
     ' Return function
     LoadLineControls = lRtn
     Exit Function
errHandler:
     MsgBox "Error:  " & Err.Number + ".  " + Err.Description, vbOKOnly + _
          vbApplicationModal + vbExclamation, "MixerSpy Error"
End Function ' LoadLineControls

'******************************************************************************
' LoadMixerCaps - loads mixer class properties into listview when a mixer
' node is selected in the treeview
'******************************************************************************
Private Sub LoadMixerCaps(sKey As String)
     Dim itmX As ListItem
     Dim lID As Long
     On Error GoTo errHandler
     If Len(sKey) > False Then
          ' change frame caption
          frameInfo.Caption = "Mixer Capabilities"
          ' clear items
          lvwItems.ListItems.Clear
          ' Fill mixercap fields
          With m_colMx(sKey)
               Set itmX = lvwItems.ListItems.Add(1, , "Product Name")
                    itmX.SubItems(1) = .ProductName
               Set itmX = lvwItems.ListItems.Add(2, , "Driver Manufacturer")
                    itmX.SubItems(1) = GetManufacturer(.ManufacturerID)
               Set itmX = lvwItems.ListItems.Add(3, , "Driver Identifier")
                    itmX.SubItems(1) = GetProductID(.ManufacturerID, .ProductID)
               Set itmX = lvwItems.ListItems.Add(4, , "Driver Version")
                    itmX.SubItems(1) = .DriverVersion
               Set itmX = lvwItems.ListItems.Add(5, , "Device ID")
                    itmX.SubItems(1) = CStr(.DeviceID)
               Set itmX = lvwItems.ListItems.Add(6, , "Destination Lines")
                    itmX.SubItems(1) = CStr(.Destinations)
               Set itmX = lvwItems.ListItems.Add(7, , "Mixer Handle")
                    itmX.SubItems(1) = CStr(.hMixer)
          End With
     End If
     ' Set control panel
     SetControlPanel False
     Set itmX = Nothing
     Exit Sub
errHandler:
     MsgBox "Error:  " & Err.Number + ".  " + Err.Description, vbOKOnly + _
          vbApplicationModal + vbExclamation, "MixerSpy Error"
End Sub ' LoadMixerCaps

'******************************************************************************
' LoadLineInfo - loads line class properties into listview when a destination
' or source line node is selected in the treeview
'******************************************************************************
Private Sub LoadLineInfo(sKey As String)
     Dim itmX As ListItem
     On Error GoTo errHandler
     If Len(sKey) > False Then
          ' change frame caption
          frameInfo.Caption = "Mixer Line Info"
          ' clear items
          lvwItems.ListItems.Clear
          ' Fill mixercap fields
          With m_colLn(sKey)
               Set itmX = lvwItems.ListItems.Add(1, , "Line Source")
                    itmX.SubItems(1) = .Name
               Set itmX = lvwItems.ListItems.Add(2, , "Short Name")
                    itmX.SubItems(1) = .ShortName
               Set itmX = lvwItems.ListItems.Add(3, , "Product Name")
                    itmX.SubItems(1) = .ProductName
               Set itmX = lvwItems.ListItems.Add(4, , "Driver Manufacturer")
                    itmX.SubItems(1) = GetManufacturer(.ManufacturerID)
               Set itmX = lvwItems.ListItems.Add(5, , "Driver Identifier")
                    itmX.SubItems(1) = GetProductID(.ManufacturerID, .ProductID)
               Set itmX = lvwItems.ListItems.Add(6, , "Driver Version")
                    itmX.SubItems(1) = .DriverVersion
               Set itmX = lvwItems.ListItems.Add(7, , "Line ID")
                    itmX.SubItems(1) = .LineID
               Set itmX = lvwItems.ListItems.Add(8, , "Component Type")
                    itmX.SubItems(1) = GetLineComponent(.ComponentType)
               Set itmX = lvwItems.ListItems.Add(9, , "Channels")
                    itmX.SubItems(1) = .Channels
               Set itmX = lvwItems.ListItems.Add(10, , "Number of Controls")
                    itmX.SubItems(1) = .Controls
               Set itmX = lvwItems.ListItems.Add(11, , "Target Type")
                    itmX.SubItems(1) = .TargetType
               Set itmX = lvwItems.ListItems.Add(12, , "Mixer Handle")
                    itmX.SubItems(1) = .hMixer
          End With
     End If
     ' Set control panel
     SetControlPanel False
     Set itmX = Nothing
     Exit Sub
errHandler:
     MsgBox "Error:  " & Err.Number + ".  " + Err.Description, vbOKOnly + _
          vbApplicationModal + vbExclamation, "MixerSpy Error"
End Sub ' LoadLineInfo

'******************************************************************************
' LoadControlInfo - loads control class properties into listview when a control
' node is selected in the treeview.  Sets the value of the controls in the
' control panel to the value specified in the class.
'******************************************************************************
Private Sub LoadControlInfo(sKey As String)
     Dim itmX As ListItem
     Dim iIdx As Integer
     Dim lLoop As Long
     Dim lVal As Long
     On Error GoTo errHandler
     ' change frame caption
     frameInfo.Caption = "Mixer Line Control Info"
     ' clear items
     lvwItems.ListItems.Clear
     ' Fill mixercap fields
     With m_colCtl(sKey)
          ' Get latest value
          m_colCtl(sKey).GetControlValue lVal
          ' Store value
          m_colCtl(sKey).Value = lVal
          ' Set treeview items
          Set itmX = lvwItems.ListItems.Add(1, , "Control ID")
               itmX.SubItems(1) = .ControlID
          Set itmX = lvwItems.ListItems.Add(2, , "Control Description")
               itmX.SubItems(1) = .ControlDesc
          Set itmX = lvwItems.ListItems.Add(3, , "Control Name")
               itmX.SubItems(1) = .ControlName
          Set itmX = lvwItems.ListItems.Add(4, , "Control Short Name")
               itmX.SubItems(1) = .ControlShortName
          Set itmX = lvwItems.ListItems.Add(5, , "Control Type")
               itmX.SubItems(1) = .ControlType
               itmX.ToolTipText = .ControlDesc
          Set itmX = lvwItems.ListItems.Add(6, , "Maximum Value")
               itmX.SubItems(1) = .Maximum
          Set itmX = lvwItems.ListItems.Add(7, , "Minimum Value")
               itmX.SubItems(1) = .Minimum
          Set itmX = lvwItems.ListItems.Add(8, , "Multiple Items")
               itmX.SubItems(1) = .MultipleItems
          Set itmX = lvwItems.ListItems.Add(9, , "Mixer Handle")
               itmX.SubItems(1) = .hMixer
          Set itmX = lvwItems.ListItems.Add(10, "Value", "Control Value")
               itmX.SubItems(1) = CStr(.Value)
          ' Show control panel
          Select Case .ControlType
               ' If a slider
               Case MIXERCONTROL_CONTROLTYPE_FADER, MIXERCONTROL_CONTROLTYPE_VOLUME, _
                         MIXERCONTROL_CONTROLTYPE_BASS, MIXERCONTROL_CONTROLTYPE_TREBLE, _
                         MIXERCONTROL_CONTROLTYPE_SLIDER
                    ' set value of panel to show
                    iIdx = 1
                    ' set info label
                    lblControlType = "Control Type:  " + .ControlName
                    ' Set slider value
                    slFader.Value = .Value
                    ' Set text value
                    txtValue = GetPctStr(.Value, .Maximum)
               ' If a 0/1 control
               Case MIXERCONTROL_CONTROLTYPE_ONOFF, MIXERCONTROL_CONTROLTYPE_BOOLEAN, _
                         MIXERCONTROL_CONTROLTYPE_LOUDNESS, MIXERCONTROL_CONTROLTYPE_BUTTON, _
                         MIXERCONTROL_CONTROLTYPE_MONO, MIXERCONTROL_CONTROLTYPE_STEREOENH, _
                         MIXERCONTROL_CONTROLTYPE_MUTE
                    ' set value of panel to shoe
                    iIdx = 2
                    ' set info label
                    lblBoolean = "Control Type:  " + .ControlName
                    ' set description label
                    lblDescription = .ControlName
                    ' set button value
                    cmdOnOff.Value = .Value
               Case Else
                    ' show empty panel
                    iIdx = False
          End Select
          ' Set control panel
          SetControlPanel iIdx
     End With
     ' Kill listitem object
     Set itmX = Nothing
     Exit Sub
errHandler:
     MsgBox "Error:  " & Err.Number + ".  " + Err.Description, vbOKOnly + _
          vbApplicationModal + vbExclamation, "MixerSpy Error"
End Sub ' LoadControlInfo

Private Sub SetValue(lVal As Long)
     Dim mx As mxControl
     Dim sKey As String
     Dim lRtn As Long
     On Error GoTo errHandler
     ' +++ Currently this version does not support setting volume on separate
     ' channels or "multiple item type" controls values are set uniformly on
     ' all channels of a control
     ' Get key from selected Node
     sKey = nodeSel.Key
     ' Get associated mixer control from collection
     Set mx = m_colCtl(sKey)
     ' Pass value to SetValue call
     lRtn = mx.SetControlValue(lVal)
     ' Set mx to nothing
     Set mx = Nothing
     Exit Sub
errHandler:
     MsgBox "Error:  " & Err.Number + ".  " + Err.Description, vbOKOnly + _
          vbApplicationModal + vbExclamation, "MixerSpy Error"
End Sub ' SetValue

'******************************************************************************
' SetControlPanel - loads the appropriate frame control based on value passed
' to the sub in iIdx.
'******************************************************************************
Private Sub SetControlPanel(iIdx As Integer)
     Dim lLoop As Long
     On Error GoTo errHandler
     ' Loop through and show appropriate frame
     For lLoop = 0 To frameCtls.Count - 1
          If lLoop = iIdx Then
               ' show frame identified in iIdx parameter
               frameCtls(lLoop).Visible = True
          Else
               ' hide other frames in array
               frameCtls(lLoop).Visible = False
          End If
     Next
     Exit Sub
errHandler:
     MsgBox "Error:  " & Err.Number + ".  " + Err.Description, vbOKOnly + _
          vbApplicationModal + vbExclamation, "MixerSpy Error"
End Sub ' SetControlPanel

