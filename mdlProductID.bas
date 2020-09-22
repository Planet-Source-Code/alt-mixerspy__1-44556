Attribute VB_Name = "mdlProductID"
Option Explicit
'**************************************************************************************************
'  Copyright © 2003, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************

' Microsoft Product ID's
Public Const MM_MIDI_MAPPER = 1                      ' Midi Mapper
Public Const MM_WAVE_MAPPER = 2                      ' Wave Mapper
Public Const MM_SNDBLST_MIDIOUT = 3                  ' Sound Blaster MIDI output port
Public Const MM_SNDBLST_MIDIIN = 4                   ' Sound Blaster MIDI input port
Public Const MM_SNDBLST_SYNTH = 5                    ' Sound Blaster internal synth
Public Const MM_SNDBLST_WAVEOUT = 6                  ' Sound Blaster waveform output
Public Const MM_SNDBLST_WAVEIN = 7                   ' Sound Blaster waveform input
Public Const MM_ADLIB = 9                            ' Ad Lib Compatible synth
Public Const MM_MPU401_MIDIOUT = 10                  ' MPU 401 compatible MIDI output port
Public Const MM_MPU401_MIDIIN = 11                   ' MPU 401 compatible MIDI input port
Public Const MM_PC_JOYSTICK = 12                     ' Joystick adapter
Public Const MM_PCSPEAKER_WAVEOUT = 13               ' PC speaker waveform output
Public Const MM_MSFT_WSS_WAVEIN = 14                 ' MS Audio Board waveform input
Public Const MM_MSFT_WSS_WAVEOUT = 15                ' MS Audio Board waveform output
Public Const MM_MSFT_WSS_FMSYNTH_STEREO = 16         ' MS Audio Board  Stereo FM synth
Public Const MM_MSFT_WSS_MIXER = 17                  ' MS Audio Board Mixer Driver
Public Const MM_MSFT_WSS_OEM_WAVEIN = 18             ' MS OEM Audio Board waveform input
Public Const MM_MSFT_WSS_OEM_WAVEOUT = 19            ' MS OEM Audio Board waveform output
Public Const MM_MSFT_WSS_OEM_FMSYNTH_STEREO = 20     ' MS OEM Audio Board Stereo FM Synth
Public Const MM_MSFT_WSS_AUX = 21                    ' MS Audio Board Aux. Port
Public Const MM_MSFT_WSS_OEM_AUX = 22                ' MS OEM Audio Aux Port
Public Const MM_MSFT_GENERIC_WAVEIN = 23             ' MS Vanilla driver waveform input
Public Const MM_MSFT_GENERIC_WAVEOUT = 24            ' MS Vanilla driver wavefrom output
Public Const MM_MSFT_GENERIC_MIDIIN = 25             ' MS Vanilla driver MIDI in
Public Const MM_MSFT_GENERIC_MIDIOUT = 26            ' MS Vanilla driver MIDI  external out
Public Const MM_MSFT_GENERIC_MIDISYNTH = 27          ' MS Vanilla driver MIDI synthesizer
Public Const MM_MSFT_GENERIC_AUX_LINE = 28           ' MS Vanilla driver aux (line in)
Public Const MM_MSFT_GENERIC_AUX_MIC = 29            ' MS Vanilla driver aux (mic)
Public Const MM_MSFT_GENERIC_AUX_CD = 30             ' MS Vanilla driver aux (CD)
Public Const MM_MSFT_WSS_OEM_MIXER = 31              ' MS OEM Audio Board Mixer Driver
Public Const MM_MSFT_MSACM = 32                      ' MS Audio Compression Manager
Public Const MM_MSFT_ACM_MSADPCM = 33                ' MS ADPCM Codec
Public Const MM_MSFT_ACM_IMAADPCM = 34               ' IMA ADPCM Codec
Public Const MM_MSFT_ACM_MSFILTER = 35               ' MS Filter
Public Const MM_MSFT_ACM_GSM610 = 36                 ' GSM 610 codec
Public Const MM_MSFT_ACM_G711 = 37                   ' G.711 codec
Public Const MM_MSFT_ACM_PCM = 38                    ' PCM converter
Public Const MM_WSS_SB16_WAVEIN = 39                 ' Sound Blaster 16 waveform input
Public Const MM_WSS_SB16_WAVEOUT = 40                ' Sound Blaster 16  waveform output
Public Const MM_WSS_SB16_MIDIIN = 41                 ' Sound Blaster 16 midi-in
Public Const MM_WSS_SB16_MIDIOUT = 42                ' Sound Blaster 16 midi out
Public Const MM_WSS_SB16_SYNTH = 43                  ' Sound Blaster 16 FM Synthesis
Public Const MM_WSS_SB16_AUX_LINE = 44               ' Sound Blaster 16 aux (line in)
Public Const MM_WSS_SB16_AUX_CD = 45                 ' Sound Blaster 16 aux (CD)
Public Const MM_WSS_SB16_MIXER = 46                  ' Sound Blaster 16 mixer device
Public Const MM_WSS_SBPRO_WAVEIN = 47                ' Sound Blaster Pro waveform input
Public Const MM_WSS_SBPRO_WAVEOUT = 48               ' Sound Blaster Pro waveform output
Public Const MM_WSS_SBPRO_MIDIIN = 49                ' Sound Blaster Pro midi in
Public Const MM_WSS_SBPRO_MIDIOUT = 50               ' Sound Blaster Pro midi out
Public Const MM_WSS_SBPRO_SYNTH = 51                 ' Sound Blaster Pro FM synthesis
Public Const MM_WSS_SBPRO_AUX_LINE = 52              ' Sound Blaster Pro aux (line in )
Public Const MM_WSS_SBPRO_AUX_CD = 53                ' Sound Blaster Pro aux (CD)
Public Const MM_WSS_SBPRO_MIXER = 54                 ' Sound Blaster Pro mixer
Public Const MM_MSFT_WSS_NT_WAVEIN = 55              ' WSS NT wave in
Public Const MM_MSFT_WSS_NT_WAVEOUT = 56             ' WSS NT wave out
Public Const MM_MSFT_WSS_NT_FMSYNTH_STEREO = 57      ' WSS NT FM synth
Public Const MM_MSFT_WSS_NT_MIXER = 58               ' WSS NT mixer
Public Const MM_MSFT_WSS_NT_AUX = 59                 ' WSS NT aux
Public Const MM_MSFT_SB16_WAVEIN = 60                ' Sound Blaster 16 waveform input
Public Const MM_MSFT_SB16_WAVEOUT = 61               ' Sound Blaster 16  waveform output
Public Const MM_MSFT_SB16_MIDIIN = 62                ' Sound Blaster 16 midi-in
Public Const MM_MSFT_SB16_MIDIOUT = 63               ' Sound Blaster 16 midi out
Public Const MM_MSFT_SB16_SYNTH = 64                 ' Sound Blaster 16 FM Synthesis
Public Const MM_MSFT_SB16_AUX_LINE = 65              ' Sound Blaster 16 aux (line in)
Public Const MM_MSFT_SB16_AUX_CD = 66                ' Sound Blaster 16 aux (CD)
Public Const MM_MSFT_SB16_MIXER = 67                 ' Sound Blaster 16 mixer device
Public Const MM_MSFT_SBPRO_WAVEIN = 68               ' Sound Blaster Pro waveform input
Public Const MM_MSFT_SBPRO_WAVEOUT = 69              ' Sound Blaster Pro waveform output
Public Const MM_MSFT_SBPRO_MIDIIN = 70               ' Sound Blaster Pro midi in
Public Const MM_MSFT_SBPRO_MIDIOUT = 71              ' Sound Blaster Pro midi out
Public Const MM_MSFT_SBPRO_SYNTH = 72                ' Sound Blaster Pro FM synthesis
Public Const MM_MSFT_SBPRO_AUX_LINE = 73             ' Sound Blaster Pro aux (line in )
Public Const MM_MSFT_SBPRO_AUX_CD = 74               ' Sound Blaster Pro aux (CD)
Public Const MM_MSFT_SBPRO_MIXER = 75                ' Sound Blaster Pro mixer
Public Const MM_MSFT_MSOPL_SYNTH = 76                ' Yamaha OPL2/OPL3 compatible FM synthesis
Public Const MM_MSFT_VMDMS_LINE_WAVEIN = 80          ' Voice Modem Serial Line Wave Input
Public Const MM_MSFT_VMDMS_LINE_WAVEOUT = 81         ' Voice Modem Serial Line Wave Output
Public Const MM_MSFT_VMDMS_HANDSET_WAVEIN = 82       ' Voice Modem Serial Handset Wave Input
Public Const MM_MSFT_VMDMS_HANDSET_WAVEOUT = 83      ' Voice Modem Serial Handset Wave Output
Public Const MM_MSFT_VMDMW_LINE_WAVEIN = 84          ' Voice Modem Wrapper Line Wave Input
Public Const MM_MSFT_VMDMW_LINE_WAVEOUT = 85         ' Voice Modem Wrapper Line Wave Output
Public Const MM_MSFT_VMDMW_HANDSET_WAVEIN = 86       ' Voice Modem Wrapper Handset Wave Input
Public Const MM_MSFT_VMDMW_HANDSET_WAVEOUT = 87      ' Voice Modem Wrapper Handset Wave Output
Public Const MM_MSFT_VMDMW_MIXER = 88                ' Voice Modem Wrapper Mixer
Public Const MM_MSFT_VMDM_GAME_WAVEOUT = 89          ' Voice Modem Game Compatible Wave Device
Public Const MM_MSFT_VMDM_GAME_WAVEIN = 90           ' Voice Modem Game Compatible Wave Device
Public Const MM_MSFT_ACM_MSNAUDIO = 91
Public Const MM_MSFT_ACM_MSG723 = 92
Public Const MM_MSFT_WDMAUDIO_WAVEOUT = 100          ' Generic id for WDM Audio drivers
Public Const MM_MSFT_WDMAUDIO_WAVEIN = 101           ' Generic id for WDM Audio drivers
Public Const MM_MSFT_WDMAUDIO_MIDIOUT = 102          ' Generic id for WDM Audio drivers
Public Const MM_MSFT_WDMAUDIO_MIDIIN = 103           ' Generic id for WDM Audio drivers
Public Const MM_MSFT_WDMAUDIO_MIXER = 104            ' Generic id for WDM Audio drivers
'MM_CREATIVE product IDs
Public Const MM_CREATIVE_SB15_WAVEIN = 1             ' SB (r) 1.5 waveform input
Public Const MM_CREATIVE_SB20_WAVEIN = 2
Public Const MM_CREATIVE_SBPRO_WAVEIN = 3
Public Const MM_CREATIVE_SBP16_WAVEIN = 4
Public Const MM_CREATIVE_PHNBLST_WAVEIN = 5
Public Const MM_CREATIVE_SB15_WAVEOUT = 101
Public Const MM_CREATIVE_SB20_WAVEOUT = 102
Public Const MM_CREATIVE_SBPRO_WAVEOUT = 103
Public Const MM_CREATIVE_SBP16_WAVEOUT = 104
Public Const MM_CREATIVE_PHNBLST_WAVEOUT = 105
Public Const MM_CREATIVE_MIDIOUT = 201               ' SB (r)
Public Const MM_CREATIVE_MIDIIN = 202                ' SB (r)
Public Const MM_CREATIVE_FMSYNTH_MONO = 301          ' SB (r)
Public Const MM_CREATIVE_FMSYNTH_STEREO = 302        ' SB Pro (r) stereo synthesizer
Public Const MM_CREATIVE_MIDI_AWE32 = 303
Public Const MM_CREATIVE_AUX_CD = 401                ' SB Pro (r) aux (CD)
Public Const MM_CREATIVE_AUX_LINE = 402              ' SB Pro (r) aux (Line in )
Public Const MM_CREATIVE_AUX_MIC = 403               ' SB Pro (r) aux (mic)
Public Const MM_CREATIVE_AUX_MASTER = 404
Public Const MM_CREATIVE_AUX_PCSPK = 405
Public Const MM_CREATIVE_AUX_WAVE = 406
Public Const MM_CREATIVE_AUX_MIDI = 407
Public Const MM_CREATIVE_SBPRO_MIXER = 408
Public Const MM_CREATIVE_SB16_MIXER = 409
'/* MM_MEDIAVISION product IDs */
' Pro Audio Spectrum
Public Const MM_MEDIAVISION_PROAUDIO = &H10
Public Const MM_PROAUD_MIDIOUT = (MM_MEDIAVISION_PROAUDIO + 1)
Public Const MM_PROAUD_MIDIIN = (MM_MEDIAVISION_PROAUDIO + 2)
Public Const MM_PROAUD_SYNTH = (MM_MEDIAVISION_PROAUDIO + 3)
Public Const MM_PROAUD_WAVEOUT = (MM_MEDIAVISION_PROAUDIO + 4)
Public Const MM_PROAUD_WAVEIN = (MM_MEDIAVISION_PROAUDIO + 5)
Public Const MM_PROAUD_MIXER = (MM_MEDIAVISION_PROAUDIO + 6)
Public Const MM_PROAUD_AUX = (MM_MEDIAVISION_PROAUDIO + 7)
' Thunder Board
Public Const MM_MEDIAVISION_THUNDER = &H20
Public Const MM_THUNDER_SYNTH = (MM_MEDIAVISION_THUNDER + 3)
Public Const MM_THUNDER_WAVEOUT = (MM_MEDIAVISION_THUNDER + 4)
Public Const MM_THUNDER_WAVEIN = (MM_MEDIAVISION_THUNDER + 5)
Public Const MM_THUNDER_AUX = (MM_MEDIAVISION_THUNDER + 7)
' Audio Port
Public Const MM_MEDIAVISION_TPORT = &H40
Public Const MM_TPORT_WAVEOUT = (MM_MEDIAVISION_TPORT + 1)
Public Const MM_TPORT_WAVEIN = (MM_MEDIAVISION_TPORT + 2)
Public Const MM_TPORT_SYNTH = (MM_MEDIAVISION_TPORT + 3)
' Pro Audio Spectrum Plus
Public Const MM_MEDIAVISION_PROAUDIO_PLUS = &H50
Public Const MM_PROAUD_PLUS_MIDIOUT = (MM_MEDIAVISION_PROAUDIO_PLUS + 1)
Public Const MM_PROAUD_PLUS_MIDIIN = (MM_MEDIAVISION_PROAUDIO_PLUS + 2)
Public Const MM_PROAUD_PLUS_SYNTH = (MM_MEDIAVISION_PROAUDIO_PLUS + 3)
Public Const MM_PROAUD_PLUS_WAVEOUT = (MM_MEDIAVISION_PROAUDIO_PLUS + 4)
Public Const MM_PROAUD_PLUS_WAVEIN = (MM_MEDIAVISION_PROAUDIO_PLUS + 5)
Public Const MM_PROAUD_PLUS_MIXER = (MM_MEDIAVISION_PROAUDIO_PLUS + 6)
Public Const MM_PROAUD_PLUS_AUX = (MM_MEDIAVISION_PROAUDIO_PLUS + 7)
' Pro Audio Spectrum 16
Public Const MM_MEDIAVISION_PROAUDIO_16 = &H60
Public Const MM_PROAUD_16_MIDIOUT = (MM_MEDIAVISION_PROAUDIO_16 + 1)
Public Const MM_PROAUD_16_MIDIIN = (MM_MEDIAVISION_PROAUDIO_16 + 2)
Public Const MM_PROAUD_16_SYNTH = (MM_MEDIAVISION_PROAUDIO_16 + 3)
Public Const MM_PROAUD_16_WAVEOUT = (MM_MEDIAVISION_PROAUDIO_16 + 4)
Public Const MM_PROAUD_16_WAVEIN = (MM_MEDIAVISION_PROAUDIO_16 + 5)
Public Const MM_PROAUD_16_MIXER = (MM_MEDIAVISION_PROAUDIO_16 + 6)
Public Const MM_PROAUD_16_AUX = (MM_MEDIAVISION_PROAUDIO_16 + 7)
' Pro Audio Studio 16
Public Const MM_MEDIAVISION_PROSTUDIO_16 = &H60
Public Const MM_STUDIO_16_MIDIOUT = (MM_MEDIAVISION_PROSTUDIO_16 + 1)
Public Const MM_STUDIO_16_MIDIIN = (MM_MEDIAVISION_PROSTUDIO_16 + 2)
Public Const MM_STUDIO_16_SYNTH = (MM_MEDIAVISION_PROSTUDIO_16 + 3)
Public Const MM_STUDIO_16_WAVEOUT = (MM_MEDIAVISION_PROSTUDIO_16 + 4)
Public Const MM_STUDIO_16_WAVEIN = (MM_MEDIAVISION_PROSTUDIO_16 + 5)
Public Const MM_STUDIO_16_MIXER = (MM_MEDIAVISION_PROSTUDIO_16 + 6)
Public Const MM_STUDIO_16_AUX = (MM_MEDIAVISION_PROSTUDIO_16 + 7)
' CDPC
Public Const MM_MEDIAVISION_CDPC = &H70
Public Const MM_CDPC_MIDIOUT = (MM_MEDIAVISION_CDPC + 1)
Public Const MM_CDPC_MIDIIN = (MM_MEDIAVISION_CDPC + 2)
Public Const MM_CDPC_SYNTH = (MM_MEDIAVISION_CDPC + 3)
Public Const MM_CDPC_WAVEOUT = (MM_MEDIAVISION_CDPC + 4)
Public Const MM_CDPC_WAVEIN = (MM_MEDIAVISION_CDPC + 5)
Public Const MM_CDPC_MIXER = (MM_MEDIAVISION_CDPC + 6)
Public Const MM_CDPC_AUX = (MM_MEDIAVISION_CDPC + 7)
' Opus MV 1208 Chipsent
Public Const MM_MEDIAVISION_OPUS1208 = &H80
Public Const MM_OPUS401_MIDIOUT = (MM_MEDIAVISION_OPUS1208 + 1)
Public Const MM_OPUS401_MIDIIN = (MM_MEDIAVISION_OPUS1208 + 2)
Public Const MM_OPUS1208_SYNTH = (MM_MEDIAVISION_OPUS1208 + 3)
Public Const MM_OPUS1208_WAVEOUT = (MM_MEDIAVISION_OPUS1208 + 4)
Public Const MM_OPUS1208_WAVEIN = (MM_MEDIAVISION_OPUS1208 + 5)
Public Const MM_OPUS1208_MIXER = (MM_MEDIAVISION_OPUS1208 + 6)
Public Const MM_OPUS1208_AUX = (MM_MEDIAVISION_OPUS1208 + 7)
'// Opus MV 1216 chipset
Public Const MM_MEDIAVISION_OPUS1216 = &H90
Public Const MM_OPUS1216_MIDIOUT = (MM_MEDIAVISION_OPUS1216 + 1)
Public Const MM_OPUS1216_MIDIIN = (MM_MEDIAVISION_OPUS1216 + 2)
Public Const MM_OPUS1216_SYNTH = (MM_MEDIAVISION_OPUS1216 + 3)
Public Const MM_OPUS1216_WAVEOUT = (MM_MEDIAVISION_OPUS1216 + 4)
Public Const MM_OPUS1216_WAVEIN = (MM_MEDIAVISION_OPUS1216 + 5)
Public Const MM_OPUS1216_MIXER = (MM_MEDIAVISION_OPUS1216 + 6)
Public Const MM_OPUS1216_AUX = (MM_MEDIAVISION_OPUS1216 + 7)
' MM_ARTISOFT product IDs
Public Const MM_ARTISOFT_SBWAVEIN = 1                     '  Artisoft sounding Board waveform input
Public Const MM_ARTISOFT_SBWAVEOUT = 2                    '  Artisoft sounding Board waveform output
' MM_IBM product IDs
Public Const MM_MMOTION_WAVEAUX = 1                       '  IBM M-Motion Auxiliary Device
Public Const MM_MMOTION_WAVEOUT = 2                       '  IBM M-Motion Waveform output
Public Const MM_MMOTION_WAVEIN = 3                        '  IBM M-Motion  Waveform Input
Public Const MM_IBM_PCMCIA_WAVEIN = 11                    '  IBM waveform input
Public Const MM_IBM_PCMCIA_WAVEOUT = 12                   '  IBM Waveform output
Public Const MM_IBM_PCMCIA_SYNTH = 13                     '  IBM Midi Synthesis
Public Const MM_IBM_PCMCIA_MIDIIN = 14                    '  IBM external MIDI in
Public Const MM_IBM_PCMCIA_MIDIOUT = 15                   '  IBM external MIDI out
Public Const MM_IBM_PCMCIA_AUX = 16                       '  IBM auxiliary control
Public Const MM_IBM_THINKPAD200 = 17
Public Const MM_IBM_MWAVE_WAVEIN = 18
Public Const MM_IBM_MWAVE_WAVEOUT = 19
Public Const MM_IBM_MWAVE_MIXER = 20
Public Const MM_IBM_MWAVE_MIDIIN = 21
Public Const MM_IBM_MWAVE_MIDIOUT = 22
Public Const MM_IBM_MWAVE_AUX = 23
Public Const MM_IBM_WC_MIDIOUT = 30
Public Const MM_IBM_WC_WAVEOUT = 31
Public Const MM_IBM_WC_MIXEROUT = 33
' MM_VOCALTEC product IDs
Public Const MM_VOCALTEC_WAVEOUT = 1
Public Const MM_VOCALTEC_WAVEIN = 2
' MM_ROLAND product IDs
Public Const MM_ROLAND_RAP10_MIDIOUT = 10                 ' MM_ROLAND_RAP10
Public Const MM_ROLAND_RAP10_MIDIIN = 11                  ' MM_ROLAND_RAP10
Public Const MM_ROLAND_RAP10_SYNTH = 12                   ' MM_ROLAND_RAP10
Public Const MM_ROLAND_RAP10_WAVEOUT = 13                 ' MM_ROLAND_RAP10
Public Const MM_ROLAND_RAP10_WAVEIN = 14                  ' MM_ROLAND_RAP10
Public Const MM_ROLAND_MPU401_MIDIOUT = 15
Public Const MM_ROLAND_MPU401_MIDIIN = 16
Public Const MM_ROLAND_SMPU_MIDIOUTA = 17
Public Const MM_ROLAND_SMPU_MIDIOUTB = 18
Public Const MM_ROLAND_SMPU_MIDIINA = 19
Public Const MM_ROLAND_SMPU_MIDIINB = 20
Public Const MM_ROLAND_SC7_MIDIOUT = 21
Public Const MM_ROLAND_SC7_MIDIIN = 22
Public Const MM_ROLAND_SERIAL_MIDIOUT = 23
Public Const MM_ROLAND_SERIAL_MIDIIN = 24
Public Const MM_ROLAND_SCP_MIDIOUT = 38
Public Const MM_ROLAND_SCP_MIDIIN = 39
Public Const MM_ROLAND_SCP_WAVEOUT = 40
Public Const MM_ROLAND_SCP_WAVEIN = 41
Public Const MM_ROLAND_SCP_MIXER = 42
Public Const MM_ROLAND_SCP_AUX = 48
' MM_DSP_SOLUTIONS product IDs
Public Const MM_DSP_SOLUTIONS_WAVEOUT = 1
Public Const MM_DSP_SOLUTIONS_WAVEIN = 2
Public Const MM_DSP_SOLUTIONS_SYNTH = 3
Public Const MM_DSP_SOLUTIONS_AUX = 4
' MM_WANGLABS product IDs
Public Const MM_WANGLABS_WAVEIN1 = 1                      '  Input audio wave on CPU board models: Exec 4010, 4030, 3450; PC 251/25c, pc 461/25s , pc 461/33c
Public Const MM_WANGLABS_WAVEOUT1 = 2
' MM_TANDY product IDs
Public Const MM_TANDY_VISWAVEIN = 1
Public Const MM_TANDY_VISWAVEOUT = 2
Public Const MM_TANDY_VISBIOSSYNTH = 3
Public Const MM_TANDY_SENS_MMAWAVEIN = 4
Public Const MM_TANDY_SENS_MMAWAVEOUT = 5
Public Const MM_TANDY_SENS_MMAMIDIIN = 6
Public Const MM_TANDY_SENS_MMAMIDIOUT = 7
Public Const MM_TANDY_SENS_VISWAVEOUT = 8
Public Const MM_TANDY_PSSJWAVEIN = 9
Public Const MM_TANDY_PSSJWAVEOUT = 10
' Intel product IDs
Public Const MM_INTELOPD_WAVEIN = 1                       '  HID2 WaveAudio Driver
Public Const MM_INTELOPD_WAVEOUT = 101                    '  HID2
Public Const MM_INTELOPD_AUX = 401                        '  HID2 for mixing
Public Const MM_INTEL_NSPMODEMLINE = 501
' MM_INTERACTIVE product IDs
Public Const MM_INTERACTIVE_WAVEIN = &H45
Public Const MM_INTERACTIVE_WAVEOUT = &H45
' MM_YAMAHA product IDs
Public Const MM_YAMAHA_GSS_SYNTH = &H1
Public Const MM_YAMAHA_GSS_WAVEOUT = &H2
Public Const MM_YAMAHA_GSS_WAVEIN = &H3
Public Const MM_YAMAHA_GSS_MIDIOUT = &H4
Public Const MM_YAMAHA_GSS_MIDIIN = &H5
Public Const MM_YAMAHA_GSS_AUX = &H6
Public Const MM_YAMAHA_SERIAL_MIDIOUT = &H7
Public Const MM_YAMAHA_SERIAL_MIDIIN = &H8
Public Const MM_YAMAHA_OPL3SA_WAVEOUT = &H10
Public Const MM_YAMAHA_OPL3SA_WAVEIN = &H11
Public Const MM_YAMAHA_OPL3SA_FMSYNTH = &H12
Public Const MM_YAMAHA_OPL3SA_YSYNTH = &H13
Public Const MM_YAMAHA_OPL3SA_MIDIOUT = &H14
Public Const MM_YAMAHA_OPL3SA_MIDIIN = &H15
Public Const MM_YAMAHA_OPL3SA_MIXER = &H17
Public Const MM_YAMAHA_OPL3SA_JOYSTICK = &H18
' MM_EVEREX product IDs
Public Const MM_EVEREX_CARRIER = &H1
' MM_ECHO product IDs
Public Const MM_ECHO_SYNTH = &H1
Public Const MM_ECHO_WAVEOUT = &H2
Public Const MM_ECHO_WAVEIN = &H3
Public Const MM_ECHO_MIDIOUT = &H4
Public Const MM_ECHO_MIDIIN = &H5
Public Const MM_ECHO_AUX = &H6
' MM_SIERRA product IDs
Public Const MM_SIERRA_ARIA_MIDIOUT = &H14
Public Const MM_SIERRA_ARIA_MIDIIN = &H15
Public Const MM_SIERRA_ARIA_SYNTH = &H16
Public Const MM_SIERRA_ARIA_WAVEOUT = &H17
Public Const MM_SIERRA_ARIA_WAVEIN = &H18
Public Const MM_SIERRA_ARIA_AUX = &H19
Public Const MM_SIERRA_ARIA_AUX2 = &H20
Public Const MM_SIERRA_QUARTET_WAVEIN = &H50
Public Const MM_SIERRA_QUARTET_WAVEOUT = &H51
Public Const MM_SIERRA_QUARTET_MIDIIN = &H52
Public Const MM_SIERRA_QUARTET_MIDIOUT = &H53
Public Const MM_SIERRA_QUARTET_SYNTH = &H54
Public Const MM_SIERRA_QUARTET_AUX_CD = &H55
Public Const MM_SIERRA_QUARTET_AUX_LINE = &H56
Public Const MM_SIERRA_QUARTET_AUX_MODEM = &H57
Public Const MM_SIERRA_QUARTET_MIXER = &H58
' MM_CAT product IDs
Public Const MM_CAT_WAVEOUT = 1
' MM_DSP_GROUP product IDs
Public Const MM_DSP_GROUP_TRUESPEECH = &H1
' MM_MELABS product IDs
Public Const MM_MELABS_MIDI2GO = &H1
' MM_ESS product IDs
Public Const MM_ESS_AMWAVEOUT = &H1
Public Const MM_ESS_AMWAVEIN = &H2
Public Const MM_ESS_AMAUX = &H3
Public Const MM_ESS_AMSYNTH = &H4
Public Const MM_ESS_AMMIDIOUT = &H5
Public Const MM_ESS_AMMIDIIN = &H6
Public Const MM_ESS_MIXER = &H7
Public Const MM_ESS_AUX_CD = &H8
Public Const MM_ESS_MPU401_MIDIOUT = &H9
Public Const MM_ESS_MPU401_MIDIIN = &HA
Public Const MM_ESS_ES488_WAVEOUT = &H10
Public Const MM_ESS_ES488_WAVEIN = &H11
Public Const MM_ESS_ES488_MIXER = &H12
Public Const MM_ESS_ES688_WAVEOUT = &H13
Public Const MM_ESS_ES688_WAVEIN = &H14
Public Const MM_ESS_ES688_MIXER = &H15
Public Const MM_ESS_ES1488_WAVEOUT = &H16
Public Const MM_ESS_ES1488_WAVEIN = &H17
Public Const MM_ESS_ES1488_MIXER = &H18
Public Const MM_ESS_ES1688_WAVEOUT = &H19
Public Const MM_ESS_ES1688_WAVEIN = &H1A
Public Const MM_ESS_ES1688_MIXER = &H1B
Public Const MM_ESS_ES1788_WAVEOUT = &H1C
Public Const MM_ESS_ES1788_WAVEIN = &H1D
Public Const MM_ESS_ES1788_MIXER = &H1E
Public Const MM_ESS_ES1888_WAVEOUT = &H1F
Public Const MM_ESS_ES1888_WAVEIN = &H20
Public Const MM_ESS_ES1888_MIXER = &H21
Public Const MM_ESS_ES1868_WAVEOUT = &H22
Public Const MM_ESS_ES1868_WAVEIN = &H23
Public Const MM_ESS_ES1868_MIXER = &H24
Public Const MM_ESS_ES1878_WAVEOUT = &H25
Public Const MM_ESS_ES1878_WAVEIN = &H26
Public Const MM_ESS_ES1878_MIXER = &H27
' product IDs
Public Const MM_EPS_FMSND = 1
' MM_TRUEVISION product IDs
Public Const MM_TRUEVISION_WAVEIN1 = 1
Public Const MM_TRUEVISION_WAVEOUT1 = 2
' MM_AZTECH product IDs
Public Const MM_AZTECH_MIDIOUT = 3
Public Const MM_AZTECH_MIDIIN = 4
Public Const MM_AZTECH_WAVEIN = 17
Public Const MM_AZTECH_WAVEOUT = 18
Public Const MM_AZTECH_FMSYNTH = 20
Public Const MM_AZTECH_MIXER = 21
Public Const MM_AZTECH_PRO16_WAVEIN = 33
Public Const MM_AZTECH_PRO16_WAVEOUT = 34
Public Const MM_AZTECH_PRO16_FMSYNTH = 38
Public Const MM_AZTECH_DSP16_WAVEIN = 65
Public Const MM_AZTECH_DSP16_WAVEOUT = 66
Public Const MM_AZTECH_DSP16_FMSYNTH = 68
Public Const MM_AZTECH_DSP16_WAVESYNTH = 70
Public Const MM_AZTECH_NOVA16_WAVEIN = 71
Public Const MM_AZTECH_NOVA16_WAVEOUT = 72
Public Const MM_AZTECH_NOVA16_MIXER = 73
Public Const MM_AZTECH_WASH16_WAVEIN = 74
Public Const MM_AZTECH_WASH16_WAVEOUT = 75
Public Const MM_AZTECH_WASH16_MIXER = 76
Public Const MM_AZTECH_AUX_CD = 401
Public Const MM_AZTECH_AUX_LINE = 402
Public Const MM_AZTECH_AUX_MIC = 403
Public Const MM_AZTECH_AUX = 404
' MM_VIDEOLOGIC product IDs
Public Const MM_VIDEOLOGIC_MSWAVEIN = 1
Public Const MM_VIDEOLOGIC_MSWAVEOUT = 2
' MM_KORG product IDs
Public Const MM_KORG_PCIF_MIDIOUT = 1
Public Const MM_KORG_PCIF_MIDIIN = 2
' MM_APT product IDs
Public Const MM_APT_ACE100CD = 1
' MM_ICS product IDs
Public Const MM_ICS_WAVEDECK_WAVEOUT = 1                  '  MS WSS compatible card and driver
Public Const MM_ICS_WAVEDECK_WAVEIN = 2
Public Const MM_ICS_WAVEDECK_MIXER = 3
Public Const MM_ICS_WAVEDECK_AUX = 4
Public Const MM_ICS_WAVEDECK_SYNTH = 5
Public Const MM_ICS_WAVEDEC_SB_WAVEOUT = 6
Public Const MM_ICS_WAVEDEC_SB_WAVEIN = 7
Public Const MM_ICS_WAVEDEC_SB_FM_MIDIOUT = 8
Public Const MM_ICS_WAVEDEC_SB_MPU401_MIDIOUT = 9
Public Const MM_ICS_WAVEDEC_SB_MPU401_MIDIIN = 10
Public Const MM_ICS_WAVEDEC_SB_MIXER = 11
Public Const MM_ICS_WAVEDEC_SB_AUX = 12
Public Const MM_ICS_2115_LITE_MIDIOUT = 13
Public Const MM_ICS_2120_LITE_MIDIOUT = 14
' MM_ITERATEDSYS product IDs
Public Const MM_ITERATEDSYS_FUFCODEC = 1
' MM_METHEUS product IDs
Public Const MM_METHEUS_ZIPPER = 1
' MM_WINNOV product IDs
Public Const MM_WINNOV_CAVIAR_WAVEIN = 1
Public Const MM_WINNOV_CAVIAR_WAVEOUT = 2
Public Const MM_WINNOV_CAVIAR_VIDC = 3
Public Const MM_WINNOV_CAVIAR_CHAMPAGNE = 4               '  Fourcc is CHAM
Public Const MM_WINNOV_CAVIAR_YUV8 = 5                    '  Fourcc is YUV8
' MM_NCR product IDs
Public Const MM_NCR_BA_WAVEIN = 1
Public Const MM_NCR_BA_WAVEOUT = 2
Public Const MM_NCR_BA_SYNTH = 3
Public Const MM_NCR_BA_AUX = 4
Public Const MM_NCR_BA_MIXER = 5
' MM_VITEC product IDs
Public Const MM_VITEC_VMAKER = 1
Public Const MM_VITEC_VMPRO = 2
' MM_MOSCOM product IDs
Public Const MM_MOSCOM_VPC2400_IN = 1                     '  Four Port Voice Processing / Voice Recognition Board
Public Const MM_MOSCOM_VPC2400_OUT = 2                    '  VPC2400
' MM_SILICONSOFT product IDs
Public Const MM_SILICONSOFT_SC1_WAVEIN = 1                '  Waveform in , high sample rate
Public Const MM_SILICONSOFT_SC1_WAVEOUT = 2               '  Waveform out , high sample rate
Public Const MM_SILICONSOFT_SC2_WAVEIN = 3                '  Waveform in 2 channels, high sample rate
Public Const MM_SILICONSOFT_SC2_WAVEOUT = 4               '  Waveform out 2 channels, high sample rate
Public Const MM_SILICONSOFT_SOUNDJR2_WAVEOUT = 5            '  Waveform out, self powered, efficient
Public Const MM_SILICONSOFT_SOUNDJR2PR_WAVEIN = 6         '  Waveform in, self powered, efficient
Public Const MM_SILICONSOFT_SOUNDJR2PR_WAVEOUT = 7        '  Waveform out 2 channels, self powered, efficient
Public Const MM_SILICONSOFT_SOUNDJR3_WAVEOUT = 8          '  Waveform in 2 channels, self powered, efficient
' MM_OLIVETTI product IDs
Public Const MM_OLIVETTI_WAVEIN = 1
Public Const MM_OLIVETTI_WAVEOUT = 2
Public Const MM_OLIVETTI_MIXER = 3
Public Const MM_OLIVETTI_AUX = 4
Public Const MM_OLIVETTI_MIDIIN = 5
Public Const MM_OLIVETTI_MIDIOUT = 6
Public Const MM_OLIVETTI_SYNTH = 7
Public Const MM_OLIVETTI_JOYSTICK = 8
Public Const MM_OLIVETTI_ACM_GSM = 9
Public Const MM_OLIVETTI_ACM_ADPCM = 10
Public Const MM_OLIVETTI_ACM_CELP = 11
Public Const MM_OLIVETTI_ACM_SBC = 12
Public Const MM_OLIVETTI_ACM_OPR = 13
' MM_IOMAGIC product IDs
' The I/O Magic Tempo is a PCMCIA Type 2 audio card featuring wave audio
' record and playback, FM synthesizer, and MIDI output.  The I/O Magic
' Tempo WaveOut device supports mono and stereo PCM playback at rates
' of 7350, 11025, 22050, and  44100 samples
Public Const MM_IOMAGIC_TEMPO_WAVEOUT = 1
Public Const MM_IOMAGIC_TEMPO_WAVEIN = 2
Public Const MM_IOMAGIC_TEMPO_SYNTH = 3
Public Const MM_IOMAGIC_TEMPO_MIDIOUT = 4
Public Const MM_IOMAGIC_TEMPO_MXDOUT = 5
Public Const MM_IOMAGIC_TEMPO_AUXOUT = 6
' MM_MATSUSHITA product IDs
Public Const MM_MATSUSHITA_WAVEIN = 1
Public Const MM_MATSUSHITA_WAVEOUT = 2
Public Const MM_MATSUSHITA_FMSYNTH_STEREO = 3
Public Const MM_MATSUSHITA_MIXER = 4
Public Const MM_MATSUSHITA_AUX = 5
' MM_NEWMEDIA product IDs
Public Const MM_NEWMEDIA_WAVJAMMER = 1                    '  WSS Compatible sound card.
' MM_LYRRUS product IDs
'  Bridge is a MIDI driver that allows the the Lyrrus G-VOX hardware to
'  communicate with Windows base transcription and sequencer applications.
'  The driver also provides a mechanism for the user to configure the system
'  to their personal playing style.
Public Const MM_LYRRUS_BRIDGE_GUITAR = 1
' MM_OPTI product IDs
Public Const MM_OPTI_M16_FMSYNTH_STEREO = &H1
Public Const MM_OPTI_M16_MIDIIN = &H2
Public Const MM_OPTI_M16_MIDIOUT = &H3
Public Const MM_OPTI_M16_WAVEIN = &H4
Public Const MM_OPTI_M16_WAVEOUT = &H5
Public Const MM_OPTI_M16_MIXER = &H6
Public Const MM_OPTI_M16_AUX = &H7
Public Const MM_OPTI_P16_FMSYNTH_STEREO = &H10
Public Const MM_OPTI_P16_MIDIIN = &H11
Public Const MM_OPTI_P16_MIDIOUT = &H12
Public Const MM_OPTI_P16_WAVEIN = &H13
Public Const MM_OPTI_P16_WAVEOUT = &H14
Public Const MM_OPTI_P16_MIXER = &H15
Public Const MM_OPTI_P16_AUX = &H16
Public Const MM_OPTI_M32_WAVEIN = &H20
Public Const MM_OPTI_M32_WAVEOUT = &H21
Public Const MM_OPTI_M32_MIDIIN = &H22
Public Const MM_OPTI_M32_MIDIOUT = &H23
Public Const MM_OPTI_M32_SYNTH_STEREO = &H24
Public Const MM_OPTI_M32_MIXER = &H25
Public Const MM_OPTI_M32_AUX = &H26
'  Product IDs for     MM_ADDX    -  ADDX
Public Const MM_ADDX_PCTV_DIGITALMIX = 1                  ' MM_ADDX_PCTV_DIGITALMIX
Public Const MM_ADDX_PCTV_WAVEIN = 2                      ' MM_ADDX_PCTV_WAVEIN
Public Const MM_ADDX_PCTV_WAVEOUT = 3                     ' MM_ADDX_PCTV_WAVEOUT
Public Const MM_ADDX_PCTV_MIXER = 4                       ' MM_ADDX_PCTV_MIXER
Public Const MM_ADDX_PCTV_AUX_CD = 5                      ' MM_ADDX_PCTV_AUX_CD
Public Const MM_ADDX_PCTV_AUX_LINE = 6                    ' MM_ADDX_PCTV_AUX_LINE
'  Product IDs for     MM_AHEAD    -  Ahead, Inc.
Public Const MM_AHEAD_MULTISOUND = 1
Public Const MM_AHEAD_SOUNDBLASTER = 2
Public Const MM_AHEAD_PROAUDIO = 3
Public Const MM_AHEAD_GENERIC = 4
'  Product IDs for     MM_AMD    -  AMD
Public Const MM_AMD_INTERWAVE_WAVEIN = 1
Public Const MM_AMD_INTERWAVE_WAVEOUT = 2
Public Const MM_AMD_INTERWAVE_SYNTH = 3
Public Const MM_AMD_INTERWAVE_MIXER1 = 4
Public Const MM_AMD_INTERWAVE_MIXER2 = 5
Public Const MM_AMD_INTERWAVE_JOYSTICK = 6
Public Const MM_AMD_INTERWAVE_EX_CD = 7
Public Const MM_AMD_INTERWAVE_MIDIIN = 8
Public Const MM_AMD_INTERWAVE_MIDIOUT = 9
Public Const MM_AMD_INTERWAVE_AUX1 = 10
Public Const MM_AMD_INTERWAVE_AUX2 = 11
Public Const MM_AMD_INTERWAVE_AUX_MIC = 12
Public Const MM_AMD_INTERWAVE_AUX_CD = 13
Public Const MM_AMD_INTERWAVE_MONO_IN = 14
Public Const MM_AMD_INTERWAVE_MONO_OUT = 15
Public Const MM_AMD_INTERWAVE_EX_TELEPHONY = 16
Public Const MM_AMD_INTERWAVE_WAVEOUT_BASE = 17
Public Const MM_AMD_INTERWAVE_WAVEOUT_TREBLE = 18
Public Const MM_AMD_INTERWAVE_STEREO_ENHANCED = 19
'  Product IDs for     MM_AST    -  AST Research Inc.
Public Const MM_AST_MODEMWAVE_WAVEIN = 13
Public Const MM_AST_MODEMWAVE_WAVEOUT = 14
'  Product IDs for     MM_BROOKTREE    -  Brooktree Corporation
Public Const MM_BTV_WAVEIN = 1          ' Brooktree PCM Wave Audio In
Public Const MM_BTV_WAVEOUT = 2         ' Brooktree PCM Wave Audio Out
Public Const MM_BTV_MIDIIN = 3          ' Brooktree MIDI In
Public Const MM_BTV_MIDIOUT = 4         ' Brooktree MIDI out
Public Const MM_BTV_MIDISYNTH = 5       ' Brooktree MIDI FM synth
Public Const MM_BTV_AUX_LINE = 6        ' Brooktree Line Input
Public Const MM_BTV_AUX_MIC = 7         ' Brooktree Microphone Input
Public Const MM_BTV_AUX_CD = 8          ' Brooktree CD Input
Public Const MM_BTV_DIGITALIN = 9       ' Brooktree PCM Wave in with subcode information
Public Const MM_BTV_DIGITALOUT = 10     ' Brooktree PCM Wave out with subcode information
Public Const MM_BTV_MIDIWAVESTREAM = 11 ' Brooktree WaveStream
Public Const MM_BTV_MIXER = 12          ' Brooktree WSS Mixer driver
'  Product IDs for     MM_CANAM    -  CANAM Computers
Public Const MM_CANAM_CBXWAVEOUT = 1
Public Const MM_CANAM_CBXWAVEIN = 2
'  Product IDs for     MM_CASIO    -  Casio Computer Co., LTD
Public Const MM_CASIO_WP150_MIDIOUT = 1  ' wp150
Public Const MM_CASIO_WP150_MIDIIN = 2
'  Product IDs for     MM_COMPAQ    -  Compaq Computer Corp.
Public Const MM_COMPAQ_BB_WAVEIN = 1
Public Const MM_COMPAQ_BB_WAVEOUT = 2
Public Const MM_COMPAQ_BB_WAVEAUX = 3
'  Product IDs for     MM_COREDYNAMICS    -  Core Dynamics
Public Const MM_COREDYNAMICS_DYNAMIXHR = 1            ' DynaMax Hi-Rez
Public Const MM_COREDYNAMICS_DYNASONIX_SYNTH = 2      ' DynaSonix
Public Const MM_COREDYNAMICS_DYNASONIX_MIDI_IN = 3
Public Const MM_COREDYNAMICS_DYNASONIX_MIDI_OUT = 4
Public Const MM_COREDYNAMICS_DYNASONIX_WAVE_IN = 5
Public Const MM_COREDYNAMICS_DYNASONIX_WAVE_OUT = 6
Public Const MM_COREDYNAMICS_DYNASONIX_AUDIO_IN = 7
Public Const MM_COREDYNAMICS_DYNASONIX_AUDIO_OUT = 8
Public Const MM_COREDYNAMICS_DYNAGRAFX_VGA = 9        ' DynaGrfx
Public Const MM_COREDYNAMICS_DYNAGRAFX_WAVE_IN = 10
Public Const MM_COREDYNAMICS_DYNAGRAFX_WAVE_OUT = 11
'  Product IDs for     MM_CRYSTAL    -  Crystal Semiconductor Corporation
Public Const MM_CRYSTAL_CS4232_WAVEIN = 1
Public Const MM_CRYSTAL_CS4232_WAVEOUT = 2
Public Const MM_CRYSTAL_CS4232_WAVEMIXER = 3
Public Const MM_CRYSTAL_CS4232_WAVEAUX_AUX1 = 4
Public Const MM_CRYSTAL_CS4232_WAVEAUX_AUX2 = 5
Public Const MM_CRYSTAL_CS4232_WAVEAUX_LINE = 6
Public Const MM_CRYSTAL_CS4232_WAVEAUX_MONO = 7
Public Const MM_CRYSTAL_CS4232_WAVEAUX_MASTER = 8
Public Const MM_CRYSTAL_CS4232_MIDIIN = 9
Public Const MM_CRYSTAL_CS4232_MIDIOUT = 10
Public Const MM_CRYSTAL_CS4232_INPUTGAIN_AUX1 = 13
Public Const MM_CRYSTAL_CS4232_INPUTGAIN_LOOP = 14
'  Product IDs for     MM_DDD    -  Danka Data Devices
Public Const MM_DDD_MIDILINK_MIDIIN = 1
Public Const MM_DDD_MIDILINK_MIDIOUT = 2
'  Product IDs for     MM_DIACOUSTICS    -  DiAcoustics, Inc.
Public Const MM_DIACOUSTICS_DRUM_ACTION = 1 ' Drum Action
'  Product IDs for     MM_DIAMONDMM    -  Diamond Multimedia
Public Const MM_DIMD_PLATFORM = 0   ' Freedom Audio
Public Const MM_DIMD_DIRSOUND = 1
Public Const MM_DIMD_VIRTMPU = 2
Public Const MM_DIMD_VIRTSB = 3
Public Const MM_DIMD_VIRTJOY = 4
Public Const MM_DIMD_WAVEIN = 5
Public Const MM_DIMD_WAVEOUT = 6
Public Const MM_DIMD_MIDIIN = 7
Public Const MM_DIMD_MIDIOUT = 8
Public Const MM_DIMD_AUX_LINE = 9
Public Const MM_DIMD_MIXER = 10
'  Product IDs for     MM_DIGITAL_AUDIO_LABS    -  Digital Audio Labs, Inc.
Public Const MM_DIGITAL_AUDIO_LABS_V8 = &H10
Public Const MM_DIGITAL_AUDIO_LABS_CPRO = &H11
'  Product IDs for     MM_DIGITAL    -  Digital Equipment Corporation
Public Const MM_DIGITAL_AV320_WAVEIN = 1   ' Digital Audio Video Compression Board
Public Const MM_DIGITAL_AV320_WAVEOUT = 2  ' Digital Audio Video Compression Board
'  Product IDs for     MM_ECS    -  Electronic Courseware Systems, Inc.
Public Const MM_ECS_AADF_MIDI_IN = 10
Public Const MM_ECS_AADF_MIDI_OUT = 11
Public Const MM_ECS_AADF_WAVE2MIDI_IN = 12
'  Product IDs for     MM_ENSONIQ    -  ENSONIQ Corporation
Public Const MM_ENSONIQ_SOUNDSCAPE = &H10 ' ENSONIQ Soundscape
Public Const MM_SOUNDSCAPE_WAVEOUT = MM_ENSONIQ_SOUNDSCAPE + 1
Public Const MM_SOUNDSCAPE_WAVEOUT_AUX = MM_ENSONIQ_SOUNDSCAPE + 2
Public Const MM_SOUNDSCAPE_WAVEIN = MM_ENSONIQ_SOUNDSCAPE + 3
Public Const MM_SOUNDSCAPE_MIDIOUT = MM_ENSONIQ_SOUNDSCAPE + 4
Public Const MM_SOUNDSCAPE_MIDIIN = MM_ENSONIQ_SOUNDSCAPE + 5
Public Const MM_SOUNDSCAPE_SYNTH = MM_ENSONIQ_SOUNDSCAPE + 6
Public Const MM_SOUNDSCAPE_MIXER = MM_ENSONIQ_SOUNDSCAPE + 7
Public Const MM_SOUNDSCAPE_AUX = MM_ENSONIQ_SOUNDSCAPE + 8
'  Product IDs for     MM_FRONTIER    -  Frontier Design Group LLC
Public Const MM_FRONTIER_WAVECENTER_MIDIIN = 1 ' WaveCenter
Public Const MM_FRONTIER_WAVECENTER_MIDIOUT = 2
Public Const MM_FRONTIER_WAVECENTER_WAVEIN = 3
Public Const MM_FRONTIER_WAVECENTER_WAVEOUT = 4
'  Product IDs for     MM_GADGETLABS    -  Gadget Labs LLC
Public Const MM_GADGETLABS_WAVE44_WAVEIN = 1
Public Const MM_GADGETLABS_WAVE44_WAVEOUT = 2
Public Const MM_GADGETLABS_WAVE42_WAVEIN = 3
Public Const MM_GADGETLABS_WAVE42_WAVEOUT = 4
Public Const MM_GADGETLABS_WAVE4_MIDIIN = 5
Public Const MM_GADGETLABS_WAVE4_MIDIOUT = 6
'  Product IDs for     MM_KAY_ELEMETRICS    -  Kay Elemetrics, Inc.
Public Const MM_KAY_ELEMETRICS_CSL = &H4300
Public Const MM_KAY_ELEMETRICS_CSL_DAT = &H4308
Public Const MM_KAY_ELEMETRICS_CSL_4CHANNEL = &H4309
'  Product IDs for     MM_LERNOUT_AND_HAUSPIE    -  Lernout & Hauspie
Public Const MM_LERNOUT_ANDHAUSPIE_LHCODECACM = 1
'  Product IDs for     MM_MPTUS    -  M.P. Technologies, Inc.
Public Const MM_MPTUS_SPWAVEOUT = 1 ' Sound Pallette
'  Product IDs for     MM_MOTU    -  Mark of the Unicorn
Public Const MM_MOTU_MTP_MIDIOUT_ALL = 100
Public Const MM_MOTU_MTP_MIDIIN_1 = 101
Public Const MM_MOTU_MTP_MIDIOUT_1 = 101
Public Const MM_MOTU_MTP_MIDIIN_2 = 102
Public Const MM_MOTU_MTP_MIDIOUT_2 = 102
Public Const MM_MOTU_MTP_MIDIIN_3 = 103
Public Const MM_MOTU_MTP_MIDIOUT_3 = 103
Public Const MM_MOTU_MTP_MIDIIN_4 = 104
Public Const MM_MOTU_MTP_MIDIOUT_4 = 104
Public Const MM_MOTU_MTP_MIDIIN_5 = 105
Public Const MM_MOTU_MTP_MIDIOUT_5 = 105
Public Const MM_MOTU_MTP_MIDIIN_6 = 106
Public Const MM_MOTU_MTP_MIDIOUT_6 = 106
Public Const MM_MOTU_MTP_MIDIIN_7 = 107
Public Const MM_MOTU_MTP_MIDIOUT_7 = 107
Public Const MM_MOTU_MTP_MIDIIN_8 = 108
Public Const MM_MOTU_MTP_MIDIOUT_8 = 108
Public Const MM_MOTU_MTPII_MIDIOUT_ALL = 200
Public Const MM_MOTU_MTPII_MIDIIN_SYNC = 200
Public Const MM_MOTU_MTPII_MIDIIN_1 = 201
Public Const MM_MOTU_MTPII_MIDIOUT_1 = 201
Public Const MM_MOTU_MTPII_MIDIIN_2 = 202
Public Const MM_MOTU_MTPII_MIDIOUT_2 = 202
Public Const MM_MOTU_MTPII_MIDIIN_3 = 203
Public Const MM_MOTU_MTPII_MIDIOUT_3 = 203
Public Const MM_MOTU_MTPII_MIDIIN_4 = 204
Public Const MM_MOTU_MTPII_MIDIOUT_4 = 204
Public Const MM_MOTU_MTPII_MIDIIN_5 = 205
Public Const MM_MOTU_MTPII_MIDIOUT_5 = 205
Public Const MM_MOTU_MTPII_MIDIIN_6 = 206
Public Const MM_MOTU_MTPII_MIDIOUT_6 = 206
Public Const MM_MOTU_MTPII_MIDIIN_7 = 207
Public Const MM_MOTU_MTPII_MIDIOUT_7 = 207
Public Const MM_MOTU_MTPII_MIDIIN_8 = 208
Public Const MM_MOTU_MTPII_MIDIOUT_8 = 208
Public Const MM_MOTU_MTPII_NET_MIDIIN_1 = 209
Public Const MM_MOTU_MTPII_NET_MIDIOUT_1 = 209
Public Const MM_MOTU_MTPII_NET_MIDIIN_2 = 210
Public Const MM_MOTU_MTPII_NET_MIDIOUT_2 = 210
Public Const MM_MOTU_MTPII_NET_MIDIIN_3 = 211
Public Const MM_MOTU_MTPII_NET_MIDIOUT_3 = 211
Public Const MM_MOTU_MTPII_NET_MIDIIN_4 = 212
Public Const MM_MOTU_MTPII_NET_MIDIOUT_4 = 212
Public Const MM_MOTU_MTPII_NET_MIDIIN_5 = 213
Public Const MM_MOTU_MTPII_NET_MIDIOUT_5 = 213
Public Const MM_MOTU_MTPII_NET_MIDIIN_6 = 214
Public Const MM_MOTU_MTPII_NET_MIDIOUT_6 = 214
Public Const MM_MOTU_MTPII_NET_MIDIIN_7 = 215
Public Const MM_MOTU_MTPII_NET_MIDIOUT_7 = 215
Public Const MM_MOTU_MTPII_NET_MIDIIN_8 = 216
Public Const MM_MOTU_MTPII_NET_MIDIOUT_8 = 216
Public Const MM_MOTU_MXP_MIDIIN_MIDIOUT_ALL = 300
Public Const MM_MOTU_MXP_MIDIIN_SYNC = 300
Public Const MM_MOTU_MXP_MIDIIN_MIDIIN_1 = 301
Public Const MM_MOTU_MXP_MIDIIN_MIDIOUT_1 = 301
Public Const MM_MOTU_MXP_MIDIIN_MIDIIN_2 = 302
Public Const MM_MOTU_MXP_MIDIIN_MIDIOUT_2 = 302
Public Const MM_MOTU_MXP_MIDIIN_MIDIIN_3 = 303
Public Const MM_MOTU_MXP_MIDIIN_MIDIOUT_3 = 303
Public Const MM_MOTU_MXP_MIDIIN_MIDIIN_4 = 304
Public Const MM_MOTU_MXP_MIDIIN_MIDIOUT_4 = 304
Public Const MM_MOTU_MXP_MIDIIN_MIDIIN_5 = 305
Public Const MM_MOTU_MXP_MIDIIN_MIDIOUT_5 = 305
Public Const MM_MOTU_MXP_MIDIIN_MIDIIN_6 = 306
Public Const MM_MOTU_MXP_MIDIIN_MIDIOUT_6 = 306
Public Const MM_MOTU_MXPMPU_MIDIOUT_ALL = 400
Public Const MM_MOTU_MXPMPU_MIDIIN_SYNC = 400
Public Const MM_MOTU_MXPMPU_MIDIIN_1 = 401
Public Const MM_MOTU_MXPMPU_MIDIOUT_1 = 401
Public Const MM_MOTU_MXPMPU_MIDIIN_2 = 402
Public Const MM_MOTU_MXPMPU_MIDIOUT_2 = 402
Public Const MM_MOTU_MXPMPU_MIDIIN_3 = 403
Public Const MM_MOTU_MXPMPU_MIDIOUT_3 = 403
Public Const MM_MOTU_MXPMPU_MIDIIN_4 = 404
Public Const MM_MOTU_MXPMPU_MIDIOUT_4 = 404
Public Const MM_MOTU_MXPMPU_MIDIIN_5 = 405
Public Const MM_MOTU_MXPMPU_MIDIOUT_5 = 405
Public Const MM_MOTU_MXPMPU_MIDIIN_6 = 406
Public Const MM_MOTU_MXPMPU_MIDIOUT_6 = 406
Public Const MM_MOTU_MXN_MIDIOUT_ALL = 500
Public Const MM_MOTU_MXN_MIDIIN_SYNC = 500
Public Const MM_MOTU_MXN_MIDIIN_1 = 501
Public Const MM_MOTU_MXN_MIDIOUT_1 = 501
Public Const MM_MOTU_MXN_MIDIIN_2 = 502
Public Const MM_MOTU_MXN_MIDIOUT_2 = 502
Public Const MM_MOTU_MXN_MIDIIN_3 = 503
Public Const MM_MOTU_MXN_MIDIOUT_3 = 503
Public Const MM_MOTU_MXN_MIDIIN_4 = 504
Public Const MM_MOTU_MXN_MIDIOUT_4 = 504
Public Const MM_MOTU_FLYER_MIDI_IN_SYNC = 600
Public Const MM_MOTU_FLYER_MIDI_IN_A = 601
Public Const MM_MOTU_FLYER_MIDI_OUT_A = 601
Public Const MM_MOTU_FLYER_MIDI_IN_B = 602
Public Const MM_MOTU_FLYER_MIDI_OUT_B = 602
Public Const MM_MOTU_PKX_MIDI_IN_SYNC = 700
Public Const MM_MOTU_PKX_MIDI_IN_A = 701
Public Const MM_MOTU_PKX_MIDI_OUT_A = 701
Public Const MM_MOTU_PKX_MIDI_IN_B = 702
Public Const MM_MOTU_PKX_MIDI_OUT_B = 702
Public Const MM_MOTU_DTX_MIDI_IN_SYNC = 800
Public Const MM_MOTU_DTX_MIDI_IN_A = 801
Public Const MM_MOTU_DTX_MIDI_OUT_A = 801
Public Const MM_MOTU_DTX_MIDI_IN_B = 802
Public Const MM_MOTU_DTX_MIDI_OUT_B = 802
Public Const MM_MOTU_MTPAV_MIDIOUT_ALL = 900
Public Const MM_MOTU_MTPAV_MIDIIN_SYNC = 900
Public Const MM_MOTU_MTPAV_MIDIIN_1 = 901
Public Const MM_MOTU_MTPAV_MIDIOUT_1 = 901
Public Const MM_MOTU_MTPAV_MIDIIN_2 = 902
Public Const MM_MOTU_MTPAV_MIDIOUT_2 = 902
Public Const MM_MOTU_MTPAV_MIDIIN_3 = 903
Public Const MM_MOTU_MTPAV_MIDIOUT_3 = 903
Public Const MM_MOTU_MTPAV_MIDIIN_4 = 904
Public Const MM_MOTU_MTPAV_MIDIOUT_4 = 904
Public Const MM_MOTU_MTPAV_MIDIIN_5 = 905
Public Const MM_MOTU_MTPAV_MIDIOUT_5 = 905
Public Const MM_MOTU_MTPAV_MIDIIN_6 = 906
Public Const MM_MOTU_MTPAV_MIDIOUT_6 = 906
Public Const MM_MOTU_MTPAV_MIDIIN_7 = 907
Public Const MM_MOTU_MTPAV_MIDIOUT_7 = 907
Public Const MM_MOTU_MTPAV_MIDIIN_8 = 908
Public Const MM_MOTU_MTPAV_MIDIOUT_8 = 908
Public Const MM_MOTU_MTPAV_NET_MIDIIN_1 = 909
Public Const MM_MOTU_MTPAV_NET_MIDIOUT_1 = 909
Public Const MM_MOTU_MTPAV_NET_MIDIIN_2 = 910
Public Const MM_MOTU_MTPAV_NET_MIDIOUT_2 = 910
Public Const MM_MOTU_MTPAV_NET_MIDIIN_3 = 911
Public Const MM_MOTU_MTPAV_NET_MIDIOUT_3 = 911
Public Const MM_MOTU_MTPAV_NET_MIDIIN_4 = 912
Public Const MM_MOTU_MTPAV_NET_MIDIOUT_4 = 912
Public Const MM_MOTU_MTPAV_NET_MIDIIN_5 = 913
Public Const MM_MOTU_MTPAV_NET_MIDIOUT_5 = 913
Public Const MM_MOTU_MTPAV_NET_MIDIIN_6 = 914
Public Const MM_MOTU_MTPAV_NET_MIDIOUT_6 = 914
Public Const MM_MOTU_MTPAV_NET_MIDIIN_7 = 915
Public Const MM_MOTU_MTPAV_NET_MIDIOUT_7 = 915
Public Const MM_MOTU_MTPAV_NET_MIDIIN_8 = 916
Public Const MM_MOTU_MTPAV_NET_MIDIOUT_8 = 916
Public Const MM_MOTU_MTPAV_MIDIIN_ADAT = 917
Public Const MM_MOTU_MTPAV_MIDIOUT_ADAT = 917
'  Product IDs for     MM_MIRO    -  miro Computer Products AG
Public Const MM_MIRO_MOVIEPRO = 1 ' miroMOVIE pro
Public Const MM_MIRO_VIDEOD1 = 2  ' miroVIDEO D1
Public Const MM_MIRO_VIDEODC1TV = 3 ' miroVIDEO DC1 tv
Public Const MM_MIRO_VIDEOTD = 4 ' miroVIDEO 10/20 TD
Public Const MM_MIRO_DC30_WAVEOUT = 5
Public Const MM_MIRO_DC30_WAVEIN = 6
Public Const MM_MIRO_DC30_MIX = 7
'  Product IDs for     MM_NEC    -  NEC
Public Const MM_NEC_73_86_SYNTH = 5
Public Const MM_NEC_73_86_WAVEOUT = 6
Public Const MM_NEC_73_86_WAVEIN = 7
Public Const MM_NEC_26_SYNTH = 9
Public Const MM_NEC_MPU401_MIDIOUT = 10
Public Const MM_NEC_MPU401_MIDIIN = 11
Public Const MM_NEC_JOYSTICK = 12
'  Product IDs for     MM_NORRIS    -  Norris Communications, Inc.
Public Const MM_NORRIS_VOICELINK = 1
'  Product IDs for     MM_NORTHERN_TELECOM    -  Northern Telecom Limited
Public Const MM_NORTEL_MPXAC_WAVEIN = 1 '  MPX Audio Card Wave Input Device
Public Const MM_NORTEL_MPXAC_WAVEOUT = 2 ' MPX Audio Card Wave Output Device
'  Product IDs for     MM_NVIDIA    -  NVidia Corporation
Public Const MM_NVIDIA_WAVEOUT = 1
Public Const MM_NVIDIA_WAVEIN = 2
Public Const MM_NVIDIA_MIDIOUT = 3
Public Const MM_NVIDIA_MIDIIN = 4
Public Const MM_NVIDIA_GAMEPORT = 5
Public Const MM_NVIDIA_MIXER = 6
Public Const MM_NVIDIA_AUX = 7
'  Product IDs for     MM_OKSORI    -  OKSORI Co., Ltd.
Public Const MM_OKSORI_BASE = 0 ' Oksori Base
Public Const MM_OKSORI_OSR8_WAVEOUT = MM_OKSORI_BASE + 1 ' Oksori 8bit Wave out
Public Const MM_OKSORI_OSR8_WAVEIN = MM_OKSORI_BASE + 2  ' Oksori 8bit Wave in
Public Const MM_OKSORI_OSR16_WAVEOUT = MM_OKSORI_BASE + 3 ' Oksori 16 bit Wave out
Public Const MM_OKSORI_OSR16_WAVEIN = MM_OKSORI_BASE + 4 ' Oksori 16 bit Wave in
Public Const MM_OKSORI_FM_OPL4 = MM_OKSORI_BASE + 5 ' Oksori FM Synth Yamaha OPL4
Public Const MM_OKSORI_MIX_MASTER = MM_OKSORI_BASE + 6 ' Oksori DSP Mixer - Master Volume
Public Const MM_OKSORI_MIX_WAVE = MM_OKSORI_BASE + 7 ' Oksori DSP Mixer - Wave Volume
Public Const MM_OKSORI_MIX_FM = MM_OKSORI_BASE + 8 ' Oksori DSP Mixer - FM Volume
Public Const MM_OKSORI_MIX_LINE = MM_OKSORI_BASE + 9 ' Oksori DSP Mixer - Line Volume
Public Const MM_OKSORI_MIX_CD = MM_OKSORI_BASE + 10 ' Oksori DSP Mixer - CD Volume
Public Const MM_OKSORI_MIX_MIC = MM_OKSORI_BASE + 11 ' Oksori DSP Mixer - MIC Volume
Public Const MM_OKSORI_MIX_ECHO = MM_OKSORI_BASE + 12 ' Oksori DSP Mixer - Echo Volume
Public Const MM_OKSORI_MIX_AUX1 = MM_OKSORI_BASE + 13 ' Oksori AD1848 - AUX1 Volume
Public Const MM_OKSORI_MIX_LINE1 = MM_OKSORI_BASE + 14 ' Oksori AD1848 - LINE1 Volume
Public Const MM_OKSORI_EXT_MIC1 = MM_OKSORI_BASE + 15 ' Oksori External - One Mic Connect
Public Const MM_OKSORI_EXT_MIC2 = MM_OKSORI_BASE + 16 ' Oksori External - Two Mic Connect
Public Const MM_OKSORI_MIDIOUT = MM_OKSORI_BASE + 17 ' Oksori MIDI Out Device
Public Const MM_OKSORI_MIDIIN = MM_OKSORI_BASE + 18 ' Oksori MIDI In Device
Public Const MM_OKSORI_MPEG_CDVISION = MM_OKSORI_BASE + 19 ' Oksori CD-Vision MPEG Decoder
'  Product IDs for     MM_OSITECH    -  Ositech Communications Inc.
Public Const MM_OSITECH_TRUMPCARD = 1                     ' Trumpcard
'  Product IDs for     MM_OSPREY    -  Osprey Technologies, Inc.
Public Const MM_OSPREY_1000WAVEIN = 1
Public Const MM_OSPREY_1000WAVEOUT = 2
'  Product IDs for     MM_QUARTERDECK    -  Quarterdeck Corporation
Public Const MM_QUARTERDECK_LHWAVEIN = 0 ' Quarterdeck L&H Codec Wave In
Public Const MM_QUARTERDECK_LHWAVEOUT = 1 ' Quarterdeck L&H Codec Wave Out
'  Product IDs for     MM_RHETOREX    -  Rhetorex Inc
Public Const MM_RHETOREX_WAVEIN = 1
Public Const MM_RHETOREX_WAVEOUT = 2
'  Product IDs for     MM_ROCKWELL    -  Rockwell International
Public Const MM_VOICEMIXER = 1
Public Const ROCKWELL_WA1_WAVEIN = 100
Public Const ROCKWELL_WA1_WAVEOUT = 101
Public Const ROCKWELL_WA1_SYNTH = 102
Public Const ROCKWELL_WA1_MIXER = 103
Public Const ROCKWELL_WA1_MPU401_IN = 104
Public Const ROCKWELL_WA1_MPU401_OUT = 105
Public Const ROCKWELL_WA2_WAVEIN = 200
Public Const ROCKWELL_WA2_WAVEOUT = 201
Public Const ROCKWELL_WA2_SYNTH = 202
Public Const ROCKWELL_WA2_MIXER = 203
Public Const ROCKWELL_WA2_MPU401_IN = 204
Public Const ROCKWELL_WA2_MPU401_OUT = 205
'  Product IDs for     MM_S3    -  S3
Public Const MM_S3_WAVEOUT = &H1
Public Const MM_S3_WAVEIN = &H2
Public Const MM_S3_MIDIOUT = &H3
Public Const MM_S3_MIDIIN = &H4
Public Const MM_S3_FMSYNTH = &H5
Public Const MM_S3_MIXER = &H6
Public Const MM_S3_AUX = &H7
'  Product IDs for     MM_SEERSYS    -  Seer Systems, Inc.
Public Const MM_SEERSYS_SEERSYNTH = 1
Public Const MM_SEERSYS_SEERWAVE = 2
Public Const MM_SEERSYS_SEERMIX = 3
'  Product IDs for     MM_SOFTSOUND    -  Softsound, Ltd.
Public Const MM_SOFTSOUND_CODEC = 1
'  Product IDs for     MM_SOUNDESIGNS    -  SounDesignS M.C.S. Ltd.
Public Const MM_SOUNDESIGNS_WAVEIN = 1
Public Const MM_SOUNDESIGNS_WAVEOUT = 2
'  Product IDs for     MM_SPECTRUM_SIGNAL_PROCESSING    -  Spectrum Signal Processing, Inc.
Public Const MM_SSP_SNDFESWAVEIN = 1                      ' Sound Festa Wave In Device
Public Const MM_SSP_SNDFESWAVEOUT = 2                     ' Sound Festa Wave Out Device
Public Const MM_SSP_SNDFESMIDIIN = 3                      ' Sound Festa MIDI In Device
Public Const MM_SSP_SNDFESMIDIOUT = 4                     ' Sound Festa MIDI Out Device
Public Const MM_SSP_SNDFESSYNTH = 5                       ' Sound Festa MIDI Synth Device
Public Const MM_SSP_SNDFESMIX = 6                         ' Sound Festa Mixer Device
Public Const MM_SSP_SNDFESAUX = 7                         ' Sound Festa Auxilliary Device
'  Product IDs for     MM_TDK    -  TDK Corporation
Public Const MM_TDK_MW_MIDI_SYNTH = 1
Public Const MM_TDK_MW_MIDI_IN = 2
Public Const MM_TDK_MW_MIDI_OUT = 3
Public Const MM_TDK_MW_WAVE_IN = 4
Public Const MM_TDK_MW_WAVE_OUT = 5
Public Const MM_TDK_MW_AUX = 6
Public Const MM_TDK_MW_MIXER = 10
Public Const MM_TDK_MW_AUX_MASTER = 100
Public Const MM_TDK_MW_AUX_BASS = 101
Public Const MM_TDK_MW_AUX_TREBLE = 102
Public Const MM_TDK_MW_AUX_MIDI_VOL = 103
Public Const MM_TDK_MW_AUX_WAVE_VOL = 104
Public Const MM_TDK_MW_AUX_WAVE_RVB = 105
Public Const MM_TDK_MW_AUX_WAVE_CHR = 106
Public Const MM_TDK_MW_AUX_VOL = 107
Public Const MM_TDK_MW_AUX_RVB = 108
Public Const MM_TDK_MW_AUX_CHR = 109
'  Product IDs for     MM_TURTLE_BEACH    -  Turtle Beach, Inc.
Public Const MM_TBS_TROPEZ_WAVEIN = 37
Public Const MM_TBS_TROPEZ_WAVEOUT = 38
Public Const MM_TBS_TROPEZ_AUX1 = 39
Public Const MM_TBS_TROPEZ_AUX2 = 40
Public Const MM_TBS_TROPEZ_LINE = 41
'  Product IDs for     MM_VIENNASYS    -  Vienna Systems
Public Const MM_VIENNASYS_TSP_WAVE_DRIVER = 1
'  Product IDs for     MM_VIONA    -  Viona Development GmbH
Public Const MM_VIONA_QVINPCI_MIXER = 1                   ' Q-Motion PCI II/Bravado 2000
Public Const MM_VIONA_QVINPCI_WAVEIN = 2
Public Const MM_VIONA_QVINPCI_WAVEOUT = 3
Public Const MM_VIONA_BUSTER_MIXER = 4                    ' Buster
Public Const MM_VIONA_CINEMASTER_MIXER = 5                ' Cinemaster
Public Const MM_VIONA_CONCERTO_MIXER = 6                  ' Concerto
'  Product IDs for     MM_WILDCAT    -  Wildcat Canyon Software
Public Const MM_WILDCAT_AUTOSCOREMIDIIN = 1               ' Autoscore
'  Product IDs for     MM_WILLOWPOND    -  Willow Pond Corporation
Public Const MM_WILLOWPOND_FMSYNTH_STEREO = 20
Public Const MM_WILLOWPOND_SNDPORT_WAVEIN = 100
Public Const MM_WILLOWPOND_SNDPORT_WAVEOUT = 101
Public Const MM_WILLOWPOND_SNDPORT_MIXER = 102
Public Const MM_WILLOWPOND_SNDPORT_AUX = 103
Public Const MM_WILLOWPOND_PH_WAVEIN = 104
Public Const MM_WILLOWPOND_PH_WAVEOUT = 105
Public Const MM_WILLOWPOND_PH_MIXER = 106
Public Const MM_WILLOWPOND_PH_AUX = 107
'  Product IDs for     MM_WORKBIT    -  Workbit Corporation
Public Const MM_WORKBIT_MIXER = 1                        ' Harmony Mixer
Public Const MM_WORKBIT_WAVEOUT = 2                      ' Harmony Mixer
Public Const MM_WORKBIT_WAVEIN = 3                       ' Harmony Mixer
Public Const MM_WORKBIT_MIDIIN = 4                       ' Harmony Mixer
Public Const MM_WORKBIT_MIDIOUT = 5                      ' Harmony Mixer
Public Const MM_WORKBIT_FMSYNTH = 6                      ' Harmony Mixer
Public Const MM_WORKBIT_AUX = 7                          ' Harmony Mixer
Public Const MM_WORKBIT_JOYSTICK = 8
'  Product IDs for     MM_FRAUNHOFER_IIS -  Fraunhofer
Public Const MM_FHGIIS_MPEGLAYER3 = 10


Public Function GetProductID(ByVal lManID As Long, ByVal lID As Long) As String
     Dim sStr As String
     Select Case lManID
          Case MM_MICROSOFT
               Select Case lID
                    Case MM_MIDI_MAPPER
                         sStr = "MM_MIDI_MAPPER"
                    Case MM_WAVE_MAPPER
                         sStr = "MM_WAVE_MAPPER"
                    Case MM_SNDBLST_MIDIOUT
                         sStr = "MM_SNDBLST_MIDIOUT"
                    Case MM_SNDBLST_MIDIIN
                         sStr = "MM_SNDBLST_MIDIIN"
                    Case MM_SNDBLST_SYNTH
                         sStr = "MM_SNDBLST_SYNTH"
                    Case MM_SNDBLST_WAVEOUT
                         sStr = "MM_SNDBLST_WAVEOUT"
                    Case MM_SNDBLST_WAVEIN
                         sStr = "MM_SNDBLST_WAVEIN"
                    Case MM_ADLIB
                         sStr = "MM_ADLIB"
                    Case MM_MPU401_MIDIOUT
                         sStr = "MM_MPU401_MIDIOUT"
                    Case MM_MPU401_MIDIIN
                         sStr = "MM_MPU401_MIDIIN"
                    Case MM_PC_JOYSTICK
                         sStr = "MM_PC_JOYSTICK"
                    Case MM_PCSPEAKER_WAVEOUT
                         sStr = "MM_PCSPEAKER_WAVEOUT"
                    Case MM_MSFT_WSS_WAVEIN
                         sStr = "MM_MSFT_WSS_WAVEIN"
                    Case MM_MSFT_WSS_WAVEOUT
                         sStr = "MM_MSFT_WSS_WAVEOUT"
                    Case MM_MSFT_WSS_FMSYNTH_STEREO
                         sStr = "MM_MSFT_WSS_FMSYNTH_STEREO"
                    Case MM_MSFT_WSS_MIXER
                         sStr = "MM_MSFT_WSS_MIXER"
                    Case MM_MSFT_WSS_OEM_WAVEIN
                         sStr = "MM_MSFT_WSS_OEM_WAVEIN"
                    Case MM_MSFT_WSS_OEM_WAVEOUT
                         sStr = "MM_MSFT_WSS_OEM_WAVEOUT"
                    Case MM_MSFT_WSS_OEM_FMSYNTH_STEREO
                         sStr = "MM_MSFT_WSS_OEM_FMSYNTH_STEREO"
                    Case MM_MSFT_WSS_AUX
                         sStr = "MM_MSFT_WSS_AUX"
                    Case MM_MSFT_WSS_OEM_AUX
                         sStr = "MM_MSFT_WSS_OEM_AUX"
                    Case MM_MSFT_GENERIC_WAVEIN
                         sStr = "MM_MSFT_GENERIC_WAVEIN"
                    Case MM_MSFT_GENERIC_WAVEOUT
                         sStr = "MM_MSFT_GENERIC_WAVEOUT"
                    Case MM_MSFT_GENERIC_MIDIIN
                         sStr = "MM_MSFT_GENERIC_MIDIIN"
                    Case MM_MSFT_GENERIC_MIDIOUT
                         sStr = "MM_MSFT_GENERIC_MIDIOUT"
                    Case MM_MSFT_GENERIC_MIDISYNTH
                         sStr = "MM_MSFT_GENERIC_MIDISYNTH"
                    Case MM_MSFT_GENERIC_AUX_LINE
                         sStr = "MM_MSFT_GENERIC_AUX_LINE"
                    Case MM_MSFT_GENERIC_AUX_MIC
                         sStr = "MM_MSFT_GENERIC_AUX_MIC"
                    Case MM_MSFT_GENERIC_AUX_CD
                         sStr = "MM_MSFT_GENERIC_AUX_CD"
                    Case MM_MSFT_WSS_OEM_MIXER
                         sStr = "MM_MSFT_WSS_OEM_MIXER"
                    Case MM_MSFT_MSACM
                         sStr = "MM_MSFT_MSACM"
                    Case MM_MSFT_ACM_MSADPCM
                         sStr = "MM_MSFT_ACM_MSADPCM"
                    Case MM_MSFT_ACM_IMAADPCM
                         sStr = "MM_MSFT_ACM_IMAADPCM"
                    Case MM_MSFT_ACM_MSFILTER
                         sStr = "MM_MSFT_ACM_MSFILTER"
                    Case MM_MSFT_ACM_GSM610
                         sStr = "MM_MSFT_ACM_GSM610"
                    Case MM_MSFT_ACM_G711
                         sStr = "MM_MSFT_ACM_G711"
                    Case MM_MSFT_ACM_PCM
                         sStr = "MM_MSFT_ACM_PCM"
                    Case MM_WSS_SB16_WAVEIN
                         sStr = "MM_WSS_SB16_WAVEIN"
                    Case MM_WSS_SB16_WAVEOUT
                         sStr = "MM_WSS_SB16_WAVEOUT"
                    Case MM_WSS_SB16_MIDIIN
                         sStr = "MM_WSS_SB16_MIDIIN"
                    Case MM_WSS_SB16_MIDIOUT
                         sStr = "MM_WSS_SB16_MIDIOUT"
                    Case MM_WSS_SB16_SYNTH
                         sStr = "MM_WSS_SB16_SYNTH"
                    Case MM_WSS_SB16_AUX_LINE
                         sStr = "MM_WSS_SB16_AUX_LINE"
                    Case MM_WSS_SB16_AUX_CD
                         sStr = "MM_WSS_SB16_AUX_CD"
                    Case MM_WSS_SB16_MIXER
                         sStr = "MM_WSS_SB16_MIXER"
                    Case MM_WSS_SBPRO_WAVEIN
                         sStr = "MM_WSS_SBPRO_WAVEIN"
                    Case MM_WSS_SBPRO_WAVEOUT
                         sStr = "MM_WSS_SBPRO_WAVEOUT"
                    Case MM_WSS_SBPRO_MIDIIN
                         sStr = "MM_WSS_SBPRO_MIDIIN"
                    Case MM_WSS_SBPRO_MIDIOUT
                         sStr = "MM_WSS_SBPRO_MIDIOUT"
                    Case MM_WSS_SBPRO_SYNTH
                         sStr = "MM_WSS_SBPRO_SYNTH"
                    Case MM_WSS_SBPRO_AUX_LINE
                         sStr = "MM_WSS_SBPRO_AUX_LINE"
                    Case MM_WSS_SBPRO_AUX_CD
                         sStr = "MM_WSS_SBPRO_AUX_CD"
                    Case MM_WSS_SBPRO_MIXER
                         sStr = "MM_WSS_SBPRO_MIXER"
                    Case MM_MSFT_WSS_NT_WAVEIN
                         sStr = "MM_MSFT_WSS_NT_WAVEIN"
                    Case MM_MSFT_WSS_NT_WAVEOUT
                         sStr = "MM_MSFT_WSS_NT_WAVEOUT"
                    Case MM_MSFT_WSS_NT_FMSYNTH_STEREO
                         sStr = "MM_MSFT_WSS_NT_FMSYNTH_STEREO"
                    Case MM_MSFT_WSS_NT_MIXER
                         sStr = "MM_MSFT_WSS_NT_MIXER"
                    Case MM_MSFT_WSS_NT_AUX
                         sStr = "MM_MSFT_WSS_NT_AUX"
                    Case MM_MSFT_SB16_WAVEIN
                         sStr = "MM_MSFT_SB16_WAVEIN"
                    Case MM_MSFT_SB16_WAVEOUT
                         sStr = "MM_MSFT_SB16_WAVEOUT"
                    Case MM_MSFT_SB16_MIDIIN
                         sStr = "MM_MSFT_SB16_MIDIIN"
                    Case MM_MSFT_SB16_MIDIOUT
                         sStr = "MM_MSFT_SB16_MIDIOUT"
                    Case MM_MSFT_SB16_SYNTH
                         sStr = "MM_MSFT_SB16_SYNTH"
                    Case MM_MSFT_SB16_AUX_LINE
                         sStr = "MM_MSFT_SB16_AUX_LINE"
                    Case MM_MSFT_SB16_AUX_CD
                         sStr = "MM_MSFT_SB16_AUX_CD"
                    Case MM_MSFT_SB16_MIXER
                         sStr = "MM_MSFT_SB16_MIXER"
                    Case MM_MSFT_SBPRO_WAVEIN
                         sStr = "MM_MSFT_SBPRO_WAVEIN"
                    Case MM_MSFT_SBPRO_WAVEOUT
                         sStr = "MM_MSFT_SBPRO_WAVEOUT"
                    Case MM_MSFT_SBPRO_MIDIIN
                         sStr = "MM_MSFT_SBPRO_MIDIIN"
                    Case MM_MSFT_SBPRO_MIDIOUT
                         sStr = "MM_MSFT_SBPRO_MIDIOUT"
                    Case MM_MSFT_SBPRO_SYNTH
                         sStr = "MM_MSFT_SBPRO_SYNTH"
                    Case MM_MSFT_SBPRO_AUX_LINE
                         sStr = "MM_MSFT_SBPRO_AUX_LINE"
                    Case MM_MSFT_SBPRO_AUX_CD
                         sStr = "MM_MSFT_SBPRO_AUX_CD"
                    Case MM_MSFT_SBPRO_MIXER
                         sStr = "MM_MSFT_SBPRO_MIXER"
                    Case MM_MSFT_MSOPL_SYNTH
                         sStr = "MM_MSFT_MSOPL_SYNTH"
                    Case MM_MSFT_VMDMS_LINE_WAVEIN
                         sStr = "MM_MSFT_VMDMS_LINE_WAVEIN"
                    Case MM_MSFT_VMDMS_LINE_WAVEOUT
                         sStr = "MM_MSFT_VMDMS_LINE_WAVEOUT"
                    Case MM_MSFT_VMDMS_HANDSET_WAVEIN
                         sStr = "MM_MSFT_VMDMS_HANDSET_WAVEIN"
                    Case MM_MSFT_VMDMS_HANDSET_WAVEOUT
                         sStr = "MM_MSFT_VMDMS_HANDSET_WAVEOUT"
                    Case MM_MSFT_VMDMW_LINE_WAVEIN
                         sStr = "MM_MSFT_VMDMW_LINE_WAVEIN"
                    Case MM_MSFT_VMDMW_LINE_WAVEOUT
                         sStr = "MM_MSFT_VMDMW_LINE_WAVEOUT"
                    Case MM_MSFT_VMDMW_HANDSET_WAVEIN
                         sStr = "MM_MSFT_VMDMW_HANDSET_WAVEIN"
                    Case MM_MSFT_VMDMW_HANDSET_WAVEOUT
                         sStr = "MM_MSFT_VMDMW_HANDSET_WAVEOUT"
                    Case MM_MSFT_VMDMW_MIXER
                         sStr = "MM_MSFT_VMDMW_MIXER"
                    Case MM_MSFT_VMDM_GAME_WAVEOUT
                         sStr = "MM_MSFT_VMDM_GAME_WAVEOUT"
                    Case MM_MSFT_VMDM_GAME_WAVEIN
                         sStr = "MM_MSFT_VMDM_GAME_WAVEIN"
                    Case MM_MSFT_ACM_MSNAUDIO
                         sStr = "MM_MSFT_ACM_MSNAUDIO"
                    Case MM_MSFT_ACM_MSG723
                         sStr = "MM_MSFT_ACM_MSG723"
                    Case MM_MSFT_WDMAUDIO_WAVEOUT
                         sStr = "MM_MSFT_WDMAUDIO_WAVEOUT"
                    Case MM_MSFT_WDMAUDIO_WAVEIN
                         sStr = "MM_MSFT_WDMAUDIO_WAVEIN"
                    Case MM_MSFT_WDMAUDIO_MIDIOUT
                         sStr = "MM_MSFT_WDMAUDIO_MIDIOUT"
                    Case MM_MSFT_WDMAUDIO_MIDIIN
                         sStr = "MM_MSFT_WDMAUDIO_MIDIIN"
                    Case MM_MSFT_WDMAUDIO_MIXER
                         sStr = "MM_MSFT_WDMAUDIO_MIXER"
               End Select
          ' Creative Labs
          Case MM_CREATIVE
               Select Case lID
                    Case MM_CREATIVE_SB15_WAVEIN
                         sStr = "MM_CREATIVE_SB15_WAVEIN"
                    Case MM_CREATIVE_SB20_WAVEIN
                         sStr = "MM_CREATIVE_SB20_WAVEIN"
                    Case MM_CREATIVE_SBPRO_WAVEIN
                         sStr = "MM_CREATIVE_SBPRO_WAVEIN"
                    Case MM_CREATIVE_SBP16_WAVEIN
                         sStr = "MM_CREATIVE_SBP16_WAVEIN"
                    Case MM_CREATIVE_PHNBLST_WAVEIN
                         sStr = "MM_CREATIVE_PHNBLST_WAVEIN"
                    Case MM_CREATIVE_SB15_WAVEOUT
                         sStr = "MM_CREATIVE_SB15_WAVEOUT"
                    Case MM_CREATIVE_SB20_WAVEOUT
                         sStr = "MM_CREATIVE_SB20_WAVEOUT"
                    Case MM_CREATIVE_SBPRO_WAVEOUT
                         sStr = "MM_CREATIVE_SBPRO_WAVEOUT"
                    Case MM_CREATIVE_SBP16_WAVEOUT
                         sStr = "MM_CREATIVE_SBP16_WAVEOUT"
                    Case MM_CREATIVE_PHNBLST_WAVEOUT
                         sStr = "MM_CREATIVE_PHNBLST_WAVEOUT"
                    Case MM_CREATIVE_MIDIOUT
                         sStr = "MM_CREATIVE_MIDIOUT"
                    Case MM_CREATIVE_MIDIIN
                         sStr = "MM_CREATIVE_MIDIIN"
                    Case MM_CREATIVE_FMSYNTH_MONO
                         sStr = "MM_CREATIVE_FMSYNTH_MONO"
                    Case MM_CREATIVE_FMSYNTH_STEREO
                         sStr = "MM_CREATIVE_FMSYNTH_STEREO"
                    Case MM_CREATIVE_MIDI_AWE32
                         sStr = "MM_CREATIVE_MIDI_AWE32"
                    Case MM_CREATIVE_AUX_CD
                         sStr = "MM_CREATIVE_AUX_CD"
                    Case MM_CREATIVE_AUX_LINE
                         sStr = "MM_CREATIVE_AUX_LINE"
                    Case MM_CREATIVE_AUX_MIC
                         sStr = "MM_CREATIVE_AUX_MIC"
                    Case MM_CREATIVE_AUX_MASTER
                         sStr = "MM_CREATIVE_AUX_MASTER"
                    Case MM_CREATIVE_AUX_PCSPK
                         sStr = "MM_CREATIVE_AUX_PCSPK"
                    Case MM_CREATIVE_AUX_WAVE
                         sStr = "MM_CREATIVE_AUX_WAVE"
                    Case MM_CREATIVE_AUX_MIDI
                         sStr = "MM_CREATIVE_AUX_MIDI"
                    Case MM_CREATIVE_SBPRO_MIXER
                         sStr = "MM_CREATIVE_SBPRO_MIXER"
                    Case MM_CREATIVE_SB16_MIXER
                         sStr = "MM_CREATIVE_SB16_MIXER"
               End Select
          Case MM_MEDIAVISION ' Media Vision, Inc.
               Select Case lID
                    Case MM_PROAUD_MIDIOUT
                         sStr = "MM_PROAUD_MIDIOUT"
                    Case MM_PROAUD_MIDIIN
                         sStr = "MM_PROAUD_MIDIIN"
                    Case MM_PROAUD_SYNTH
                         sStr = "MM_PROAUD_SYNTH"
                    Case MM_PROAUD_WAVEOUT
                         sStr = "MM_PROAUD_WAVEOUT"
                    Case MM_PROAUD_WAVEIN
                         sStr = "MM_PROAUD_WAVEIN"
                    Case MM_PROAUD_MIXER
                         sStr = "MM_PROAUD_MIXER"
                    Case MM_PROAUD_AUX
                         sStr = "MM_PROAUD_AUX"
                    ' Thunder Board
                    Case MM_THUNDER_SYNTH
                         sStr = "MM_THUNDER_SYNTH"
                    Case MM_THUNDER_WAVEOUT
                         sStr = "MM_THUNDER_WAVEOUT"
                    Case MM_THUNDER_WAVEIN
                         sStr = "MM_THUNDER_WAVEIN"
                    Case MM_THUNDER_AUX
                         sStr = "MM_THUNDER_AUX"
                    ' Audio Port
                    Case MM_TPORT_WAVEOUT
                         sStr = "MM_TPORT_WAVEOUT"
                    Case MM_TPORT_WAVEIN
                         sStr = "MM_TPORT_WAVEIN"
                    Case MM_TPORT_SYNTH
                         sStr = "MM_TPORT_SYNTH"
                    ' Pro Audio Spectrum Plus
                    Case MM_PROAUD_PLUS_MIDIOUT
                         sStr = "MM_PROAUD_PLUS_MIDIOUT"
                    Case MM_PROAUD_PLUS_MIDIIN
                         sStr = "MM_PROAUD_PLUS_MIDIIN"
                    Case MM_PROAUD_PLUS_SYNTH
                         sStr = "MM_PROAUD_PLUS_SYNTH"
                    Case MM_PROAUD_PLUS_WAVEOUT
                         sStr = "MM_PROAUD_PLUS_WAVEOUT"
                    Case MM_PROAUD_PLUS_WAVEIN
                         sStr = "MM_PROAUD_PLUS_WAVEIN"
                    Case MM_PROAUD_PLUS_MIXER
                         sStr = "MM_PROAUD_PLUS_MIXER"
                    Case MM_PROAUD_PLUS_AUX
                         sStr = "MM_PROAUD_PLUS_AUX"
                    ' Pro Audio Spectrum 16
                    Case MM_PROAUD_16_MIDIOUT
                         sStr = "MM_PROAUD_16_MIDIOUT"
                    Case MM_PROAUD_16_MIDIIN
                         sStr = "MM_PROAUD_16_MIDIIN"
                    Case MM_PROAUD_16_SYNTH
                         sStr = "MM_PROAUD_16_SYNTH"
                    Case MM_PROAUD_16_WAVEOUT
                         sStr = "MM_PROAUD_16_WAVEOUT"
                    Case MM_PROAUD_16_WAVEIN
                         sStr = "MM_PROAUD_16_WAVEIN"
                    Case MM_PROAUD_16_MIXER
                         sStr = "MM_PROAUD_16_MIXER"
                    Case MM_PROAUD_16_AUX
                         sStr = "MM_PROAUD_16_AUX"
                    ' Pro Audio Studio 16
                    Case MM_STUDIO_16_MIDIOUT
                         sStr = "MM_STUDIO_16_MIDIOUT"
                    Case MM_STUDIO_16_MIDIIN
                         sStr = "MM_STUDIO_16_MIDIIN"
                    Case MM_STUDIO_16_SYNTH
                         sStr = "MM_STUDIO_16_SYNTH"
                    Case MM_STUDIO_16_WAVEOUT
                         sStr = "MM_STUDIO_16_WAVEOUT"
                    Case MM_STUDIO_16_WAVEIN
                         sStr = "MM_STUDIO_16_WAVEIN"
                    Case MM_STUDIO_16_MIXER
                         sStr = "MM_STUDIO_16_MIXER"
                    Case MM_STUDIO_16_AUX
                         sStr = "MM_STUDIO_16_AUX"
                    ' CDPC
                    Case MM_CDPC_MIDIOUT
                         sStr = "MM_CDPC_MIDIOUT"
                    Case MM_CDPC_MIDIIN
                         sStr = "MM_CDPC_MIDIIN"
                    Case MM_CDPC_SYNTH
                         sStr = "MM_CDPC_SYNTH"
                    Case MM_CDPC_WAVEOUT
                         sStr = "MM_CDPC_WAVEOUT"
                    Case MM_CDPC_WAVEIN
                         sStr = "MM_CDPC_WAVEIN"
                    Case MM_CDPC_MIXER
                         sStr = "MM_CDPC_MIXER"
                    Case MM_CDPC_AUX
                         sStr = "MM_CDPC_AUX"
                    ' Opus MV 1208 Chipsent
                    Case MM_OPUS401_MIDIOUT
                         sStr = "MM_OPUS401_MIDIOUT"
                    Case MM_OPUS401_MIDIIN
                         sStr = "MM_OPUS401_MIDIIN"
                    Case MM_OPUS1208_SYNTH
                         sStr = "MM_OPUS1208_SYNTH"
                    Case MM_OPUS1208_WAVEOUT
                         sStr = "MM_OPUS1208_WAVEOUT"
                    Case MM_OPUS1208_WAVEIN
                         sStr = "MM_OPUS1208_WAVEIN"
                    Case MM_OPUS1208_MIXER
                         sStr = "MM_OPUS1208_MIXER"
                    Case MM_OPUS1208_AUX
                         sStr = "MM_OPUS1208_AUX"
                    '// Opus MV 1216 chipset
                    Case MM_OPUS1216_MIDIOUT
                         sStr = "MM_OPUS1216_MIDIOUT"
                    Case MM_OPUS1216_MIDIIN
                         sStr = "MM_OPUS1216_MIDIIN"
                    Case MM_OPUS1216_SYNTH
                         sStr = "MM_OPUS1216_SYNTH"
                    Case MM_OPUS1216_WAVEOUT
                         sStr = "MM_OPUS1216_WAVEOUT"
                    Case MM_OPUS1216_WAVEIN
                         sStr = "MM_OPUS1216_WAVEIN"
                    Case MM_OPUS1216_MIXER
                         sStr = "MM_OPUS1216_MIXER"
                    Case MM_OPUS1216_AUX
                         sStr = "MM_OPUS1216_AUX"
               End Select
          Case MM_FUJITSU ' Fujitsu Corp.
               sStr = "No Associated Product ID"
          Case MM_ARTISOFT ' Artisoft, Inc.
               Select Case lID
                    Case MM_ARTISOFT_SBWAVEIN
                         sStr = "MM_ARTISOFT_SBWAVEIN"
                    Case MM_ARTISOFT_SBWAVEOUT
                         sStr = "MM_ARTISOFT_SBWAVEOUT"
               End Select
          Case MM_TURTLE_BEACH  ' Turtle Beach, Inc.
               Select Case lID
                    Case MM_TBS_TROPEZ_WAVEIN
                         sStr = "MM_TBS_TROPEZ_WAVEIN"
                    Case MM_TBS_TROPEZ_WAVEOUT
                         sStr = "MM_TBS_TROPEZ_WAVEOUT"
                    Case MM_TBS_TROPEZ_AUX1
                         sStr = "MM_TBS_TROPEZ_AUX1"
                    Case MM_TBS_TROPEZ_AUX2
                         sStr = "MM_TBS_TROPEZ_AUX2"
                    Case MM_TBS_TROPEZ_LINE
                         sStr = "MM_TBS_TROPEZ_LINE"
               End Select
          Case MM_IBM ' IBM Corporation
               Select Case lID
                    Case MM_MMOTION_WAVEAUX
                         sStr = "MM_MMOTION_WAVEAUX"
                    Case MM_MMOTION_WAVEOUT
                         sStr = "MM_MMOTION_WAVEOUT"
                    Case MM_MMOTION_WAVEIN
                         sStr = "MM_MMOTION_WAVEIN"
                    Case MM_IBM_PCMCIA_WAVEIN
                         sStr = "MM_IBM_PCMCIA_WAVEIN"
                    Case MM_IBM_PCMCIA_WAVEOUT
                         sStr = "MM_IBM_PCMCIA_WAVEOUT"
                    Case MM_IBM_PCMCIA_SYNTH
                         sStr = "MM_IBM_PCMCIA_SYNTH"
                    Case MM_IBM_PCMCIA_MIDIIN
                         sStr = "MM_IBM_PCMCIA_MIDIIN"
                    Case MM_IBM_PCMCIA_MIDIOUT
                         sStr = "MM_IBM_PCMCIA_MIDIOUT"
                    Case MM_IBM_PCMCIA_AUX
                         sStr = "MM_IBM_PCMCIA_AUX"
                    Case MM_IBM_THINKPAD200
                         sStr = "MM_IBM_THINKPAD200"
                    Case MM_IBM_MWAVE_WAVEIN
                         sStr = "MM_IBM_MWAVE_WAVEIN"
                    Case MM_IBM_MWAVE_WAVEOUT
                         sStr = "MM_IBM_MWAVE_WAVEOUT"
                    Case MM_IBM_MWAVE_MIXER
                         sStr = "MM_IBM_MWAVE_MIXER"
                    Case MM_IBM_MWAVE_MIDIIN
                         sStr = "MM_IBM_MWAVE_MIDIIN"
                    Case MM_IBM_MWAVE_MIDIOUT
                         sStr = "MM_IBM_MWAVE_MIDIOUT"
                    Case MM_IBM_MWAVE_AUX
                         sStr = "MM_IBM_MWAVE_AUX"
                    Case MM_IBM_WC_MIDIOUT
                         sStr = "MM_IBM_WC_MIDIOUT"
                    Case MM_IBM_WC_WAVEOUT
                         sStr = "MM_IBM_WC_WAVEOUT"
                    Case MM_IBM_WC_MIXEROUT
                         sStr = "MM_IBM_WC_MIXEROUT"
               End Select
          Case MM_VOCALTEC  ' Vocaltec LTD.
               Select Case lID
                    Case MM_VOCALTEC_WAVEOUT
                         sStr = "MM_VOCALTEC_WAVEOUT"
                    Case MM_VOCALTEC_WAVEIN
                         sStr = "MM_VOCALTEC_WAVEIN"
               End Select
          Case MM_ROLAND ' Roland
               Select Case lID
                    Case MM_ROLAND_RAP10_MIDIOUT
                         sStr = "MM_ROLAND_RAP10_MIDIOUT"
                    Case MM_ROLAND_RAP10_MIDIIN
                         sStr = "MM_ROLAND_RAP10_MIDIIN"
                    Case MM_ROLAND_RAP10_SYNTH
                         sStr = "MM_ROLAND_RAP10_SYNTH"
                    Case MM_ROLAND_RAP10_WAVEOUT
                         sStr = "MM_ROLAND_RAP10_WAVEOUT"
                    Case MM_ROLAND_RAP10_WAVEIN
                         sStr = "MM_ROLAND_RAP10_WAVEIN"
                    Case MM_ROLAND_MPU401_MIDIOUT
                         sStr = "MM_ROLAND_MPU401_MIDIOUT"
                    Case MM_ROLAND_MPU401_MIDIIN
                         sStr = "MM_ROLAND_MPU401_MIDIIN"
                    Case MM_ROLAND_SMPU_MIDIOUTA
                         sStr = "MM_ROLAND_SMPU_MIDIOUTA"
                    Case MM_ROLAND_SMPU_MIDIOUTB
                         sStr = "MM_ROLAND_SMPU_MIDIOUTB"
                    Case MM_ROLAND_SMPU_MIDIINA
                         sStr = "MM_ROLAND_SMPU_MIDIINA"
                    Case MM_ROLAND_SMPU_MIDIINB
                         sStr = "MM_ROLAND_SMPU_MIDIINB"
                    Case MM_ROLAND_SC7_MIDIOUT
                         sStr = "MM_ROLAND_SC7_MIDIOUT"
                    Case MM_ROLAND_SC7_MIDIIN
                         sStr = "MM_ROLAND_SC7_MIDIIN"
                    Case MM_ROLAND_SERIAL_MIDIOUT
                         sStr = "MM_ROLAND_SERIAL_MIDIOUT"
                    Case MM_ROLAND_SERIAL_MIDIIN
                         sStr = "MM_ROLAND_SERIAL_MIDIIN"
                    Case MM_ROLAND_SCP_MIDIOUT
                         sStr = "MM_ROLAND_SCP_MIDIOUT"
                    Case MM_ROLAND_SCP_MIDIIN
                         sStr = "MM_ROLAND_SCP_MIDIIN"
                    Case MM_ROLAND_SCP_WAVEOUT
                         sStr = "MM_ROLAND_SCP_WAVEOUT"
                    Case MM_ROLAND_SCP_WAVEIN
                         sStr = "MM_ROLAND_SCP_WAVEIN"
                    Case MM_ROLAND_SCP_MIXER
                         sStr = "MM_ROLAND_SCP_MIXER"
                    Case MM_ROLAND_SCP_AUX
                         sStr = "MM_ROLAND_SCP_AUX"
               End Select
          Case MM_DSP_SOLUTIONS ' DSP Solutions, Inc.
               Select Case lID
                    Case MM_DSP_SOLUTIONS_WAVEOUT
                         sStr = "MM_DSP_SOLUTIONS_WAVEOUT"
                    Case MM_DSP_SOLUTIONS_WAVEIN
                         sStr = "MM_DSP_SOLUTIONS_WAVEIN"
                    Case MM_DSP_SOLUTIONS_SYNTH
                         sStr = "MM_DSP_SOLUTIONS_SYNTH"
                    Case MM_DSP_SOLUTIONS_AUX
                         sStr = "MM_DSP_SOLUTIONS_AUX"
               End Select
          Case MM_NEC ' NEC
               Select Case lID
                    Case MM_NEC_73_86_SYNTH
                         sStr = "MM_NEC_73_86_SYNTH"
                    Case MM_NEC_73_86_WAVEOUT
                         sStr = "MM_NEC_73_86_WAVEOUT"
                    Case MM_NEC_73_86_WAVEIN
                         sStr = "MM_NEC_73_86_WAVEIN"
                    Case MM_NEC_26_SYNTH
                         sStr = "MM_NEC_26_SYNTH"
                    Case MM_NEC_MPU401_MIDIOUT
                         sStr = "MM_NEC_MPU401_MIDIOUT"
                    Case MM_NEC_MPU401_MIDIIN
                         sStr = "MM_NEC_MPU401_MIDIIN"
                    Case MM_NEC_JOYSTICK
                         sStr = "MM_NEC_JOYSTICK"
               End Select
          Case MM_ATI ' ATI
               sStr = "No Associated Product ID"
          Case MM_WANGLABS ' Wang Laboratories, Inc
               Select Case lID
                    Case MM_WANGLABS_WAVEIN1
                         sStr = "MM_WANGLABS_WAVEIN1"
                    Case MM_WANGLABS_WAVEOUT1
                         sStr = "MM_WANGLABS_WAVEOUT1"
               End Select
          Case MM_TANDY ' Tandy Corporation
               Select Case lID
                    Case MM_TANDY_VISWAVEIN
                         sStr = "MM_TANDY_VISWAVEIN"
                    Case MM_TANDY_VISWAVEOUT
                         sStr = "MM_TANDY_VISWAVEOUT"
                    Case MM_TANDY_VISBIOSSYNTH
                         sStr = "MM_TANDY_VISBIOSSYNTH"
                    Case MM_TANDY_SENS_MMAWAVEIN
                         sStr = "MM_TANDY_SENS_MMAWAVEIN"
                    Case MM_TANDY_SENS_MMAWAVEOUT
                         sStr = "MM_TANDY_SENS_MMAWAVEOUT"
                    Case MM_TANDY_SENS_MMAMIDIIN
                         sStr = "MM_TANDY_SENS_MMAMIDIIN"
                    Case MM_TANDY_SENS_MMAMIDIOUT
                         sStr = "MM_TANDY_SENS_MMAMIDIOUT"
                    Case MM_TANDY_SENS_VISWAVEOUT
                         sStr = "MM_TANDY_SENS_VISWAVEOUT"
                    Case MM_TANDY_PSSJWAVEIN
                         sStr = "MM_TANDY_PSSJWAVEIN"
                    Case MM_TANDY_PSSJWAVEOUT
                         sStr = "MM_TANDY_PSSJWAVEOUT"
               End Select
          Case MM_VOYETRA ' Voyetra
               sStr = "No Associated Product ID"
          Case MM_ANTEX ' Antex Electronics Corporation
               sStr = "No Associated Product ID"
          Case MM_ICL_PS ' ICL Personal Systems
               sStr = "No Associated Product ID"
          Case MM_INTEL ' Intel Corporation
               Select Case lID
                    Case MM_INTELOPD_WAVEIN
                         sStr = "MM_INTELOPD_WAVEIN"
                    Case MM_INTELOPD_WAVEOUT
                         sStr = "MM_INTELOPD_WAVEOUT"
                    Case MM_INTELOPD_AUX
                         sStr = "MM_INTELOPD_AUX"
                    Case MM_INTEL_NSPMODEMLINE
                         sStr = "MM_INTEL_NSPMODEMLINE"
               End Select
          Case MM_GRAVIS ' Advanced Gravis
               sStr = "No Associated Product ID"
          Case MM_VAL ' Video Associates Labs, Inc.
               sStr = "No Associated Product ID"
          Case MM_INTERACTIVE ' InterActive Inc
               Select Case lID
                    Case MM_INTERACTIVE_WAVEIN = &H45
                         sStr = "MM_INTERACTIVE_WAVEIN"
                    Case MM_INTERACTIVE_WAVEOUT = &H45
                         sStr = "MM_INTERACTIVE_WAVEOUT"
               End Select
          Case MM_YAMAHA ' Yamaha Corporation of America
               Select Case lID
                    Case MM_YAMAHA_GSS_SYNTH
                         sStr = "MM_YAMAHA_GSS_SYNTH"
                    Case MM_YAMAHA_GSS_WAVEOUT
                         sStr = "MM_YAMAHA_GSS_WAVEOUT"
                    Case MM_YAMAHA_GSS_WAVEIN
                         sStr = "MM_YAMAHA_GSS_WAVEIN"
                    Case MM_YAMAHA_GSS_MIDIOUT
                         sStr = "MM_YAMAHA_GSS_MIDIOUT"
                    Case MM_YAMAHA_GSS_MIDIIN
                         sStr = "MM_YAMAHA_GSS_MIDIIN"
                    Case MM_YAMAHA_GSS_AUX
                         sStr = "MM_YAMAHA_GSS_AUX"
                    Case MM_YAMAHA_SERIAL_MIDIOUT
                         sStr = "MM_YAMAHA_SERIAL_MIDIOUT"
                    Case MM_YAMAHA_SERIAL_MIDIIN
                         sStr = "MM_YAMAHA_SERIAL_MIDIIN"
                    Case MM_YAMAHA_OPL3SA_WAVEOUT
                         sStr = "MM_YAMAHA_OPL3SA_WAVEOUT"
                    Case MM_YAMAHA_OPL3SA_WAVEIN
                         sStr = "MM_YAMAHA_OPL3SA_WAVEIN"
                    Case MM_YAMAHA_OPL3SA_FMSYNTH
                         sStr = "MM_YAMAHA_OPL3SA_FMSYNTH"
                    Case MM_YAMAHA_OPL3SA_YSYNTH
                         sStr = "MM_YAMAHA_OPL3SA_YSYNTH"
                    Case MM_YAMAHA_OPL3SA_MIDIOUT
                         sStr = "MM_YAMAHA_OPL3SA_MIDIOUT"
                    Case MM_YAMAHA_OPL3SA_MIDIIN
                         sStr = "MM_YAMAHA_OPL3SA_MIDIIN"
                    Case MM_YAMAHA_OPL3SA_MIXER
                         sStr = "MM_YAMAHA_OPL3SA_MIXER"
                    Case MM_YAMAHA_OPL3SA_JOYSTICK
                         sStr = "MM_YAMAHA_OPL3SA_JOYSTICK"
               End Select
          Case MM_EVEREX ' Everex Systems, Inc
               sStr = "MM_EVEREX_CARRIER"
          Case MM_ECHO ' Echo Speech Corporation
               Select Case lID
                    Case MM_ECHO_SYNTH
                         sStr = "MM_ECHO_SYNTH"
                    Case MM_ECHO_WAVEOUT
                         sStr = "MM_ECHO_WAVEOUT"
                    Case MM_ECHO_WAVEIN
                         sStr = "MM_ECHO_WAVEIN"
                    Case MM_ECHO_MIDIOUT
                         sStr = "MM_ECHO_MIDIOUT"
                    Case MM_ECHO_MIDIIN
                         sStr = "MM_ECHO_MIDIIN"
                    Case MM_ECHO_AUX
                         sStr = "MM_ECHO_AUX"
               End Select
          Case MM_SIERRA ' Sierra Semiconductor Corp
               Select Case lID
                    Case MM_SIERRA_ARIA_MIDIOUT
                         sStr = "MM_SIERRA_ARIA_MIDIOUT"
                    Case MM_SIERRA_ARIA_MIDIIN
                         sStr = "MM_SIERRA_ARIA_MIDIIN"
                    Case MM_SIERRA_ARIA_SYNTH
                         sStr = "MM_SIERRA_ARIA_SYNTH"
                    Case MM_SIERRA_ARIA_WAVEOUT
                         sStr = "MM_SIERRA_ARIA_WAVEOUT"
                    Case MM_SIERRA_ARIA_WAVEIN
                         sStr = "MM_SIERRA_ARIA_WAVEIN"
                    Case MM_SIERRA_ARIA_AUX
                         sStr = "MM_SIERRA_ARIA_AUX"
                    Case MM_SIERRA_ARIA_AUX2
                         sStr = "MM_SIERRA_ARIA_AUX2"
                    Case MM_SIERRA_QUARTET_WAVEIN
                         sStr = "MM_SIERRA_QUARTET_WAVEIN"
                    Case MM_SIERRA_QUARTET_WAVEOUT
                         sStr = "MM_SIERRA_QUARTET_WAVEOUT"
                    Case MM_SIERRA_QUARTET_MIDIIN
                         sStr = "MM_SIERRA_QUARTET_MIDIIN"
                    Case MM_SIERRA_QUARTET_MIDIOUT
                         sStr = "MM_SIERRA_QUARTET_MIDIOUT"
                    Case MM_SIERRA_QUARTET_SYNTH
                         sStr = "MM_SIERRA_QUARTET_SYNTH"
                    Case MM_SIERRA_QUARTET_AUX_CD
                         sStr = "MM_SIERRA_QUARTET_AUX_CD"
                    Case MM_SIERRA_QUARTET_AUX_LINE
                         sStr = "MM_SIERRA_QUARTET_AUX_LINE"
                    Case MM_SIERRA_QUARTET_AUX_MODEM
                         sStr = "MM_SIERRA_QUARTET_AUX_MODEM"
                    Case MM_SIERRA_QUARTET_MIXER
                         sStr = "MM_SIERRA_QUARTET_MIXER"
               End Select
          Case MM_CAT ' Computer Aided Technologies
               sStr = "MM_CAT_WAVEOUT"
          Case MM_APPS ' APPS Software International
               sStr = "No Associated Product ID"
          Case MM_DSP_GROUP ' DSP Group, Inc
               sStr = "MM_DSP_GROUP_TRUESPEECH"
          Case MM_MELABS ' microEngineering Labs
               sStr = "MM_MELABS_MIDI2GO"
          Case MM_COMPUTER_FRIENDS ' Computer Friends, Inc.
               sStr = "No Associated Product ID"
          Case MM_ESS ' ESS Technology
               Select Case lID
                    Case MM_ESS_AMWAVEOUT
                         sStr = "MM_ESS_AMWAVEOUT"
                    Case MM_ESS_AMWAVEIN
                         sStr = "MM_ESS_AMWAVEIN"
                    Case MM_ESS_AMAUX
                         sStr = "MM_ESS_AMAUX"
                    Case MM_ESS_AMSYNTH
                         sStr = "MM_ESS_AMSYNTH"
                    Case MM_ESS_AMMIDIOUT
                         sStr = "MM_ESS_AMMIDIOUT"
                    Case MM_ESS_AMMIDIIN
                         sStr = "MM_ESS_AMMIDIIN"
                    Case MM_ESS_MIXER
                         sStr = "MM_ESS_MIXER"
                    Case MM_ESS_AUX_CD
                         sStr = "MM_ESS_AUX_CD"
                    Case MM_ESS_MPU401_MIDIOUT
                         sStr = "MM_ESS_MPU401_MIDIOUT"
                    Case MM_ESS_MPU401_MIDIIN
                         sStr = "MM_ESS_MPU401_MIDIIN"
                    Case MM_ESS_ES488_WAVEOUT
                         sStr = "MM_ESS_ES488_WAVEOUT"
                    Case MM_ESS_ES488_WAVEIN
                         sStr = "MM_ESS_ES488_WAVEIN"
                    Case MM_ESS_ES488_MIXER
                         sStr = "MM_ESS_ES488_MIXER"
                    Case MM_ESS_ES688_WAVEOUT
                         sStr = "MM_ESS_ES688_WAVEOUT"
                    Case MM_ESS_ES688_WAVEIN
                         sStr = "MM_ESS_ES688_WAVEIN"
                    Case MM_ESS_ES688_MIXER
                         sStr = "MM_ESS_ES688_MIXER"
                    Case MM_ESS_ES1488_WAVEOUT
                         sStr = "MM_ESS_ES1488_WAVEOUT"
                    Case MM_ESS_ES1488_WAVEIN
                         sStr = "MM_ESS_ES1488_WAVEIN"
                    Case MM_ESS_ES1488_MIXER
                         sStr = "MM_ESS_ES1488_MIXER"
                    Case MM_ESS_ES1688_WAVEOUT
                         sStr = "MM_ESS_ES1688_WAVEOUT"
                    Case MM_ESS_ES1688_WAVEIN
                         sStr = "MM_ESS_ES1688_WAVEIN"
                    Case MM_ESS_ES1688_MIXER
                         sStr = "MM_ESS_ES1688_MIXER"
                    Case MM_ESS_ES1788_WAVEOUT
                         sStr = "MM_ESS_ES1788_WAVEOUT"
                    Case MM_ESS_ES1788_WAVEIN
                         sStr = "MM_ESS_ES1788_WAVEIN"
                    Case MM_ESS_ES1788_MIXER
                         sStr = "MM_ESS_ES1788_MIXER"
                    Case MM_ESS_ES1888_WAVEOUT
                         sStr = "MM_ESS_ES1888_WAVEOUT"
                    Case MM_ESS_ES1888_WAVEIN
                         sStr = "MM_ESS_ES1888_WAVEIN"
                    Case MM_ESS_ES1888_MIXER
                         sStr = "MM_ESS_ES1888_MIXER"
                    Case MM_ESS_ES1868_WAVEOUT
                         sStr = "MM_ESS_ES1868_WAVEOUT"
                    Case MM_ESS_ES1868_WAVEIN
                         sStr = "MM_ESS_ES1868_WAVEIN"
                    Case MM_ESS_ES1868_MIXER
                         sStr = "MM_ESS_ES1868_MIXER"
                    Case MM_ESS_ES1878_WAVEOUT
                         sStr = "MM_ESS_ES1878_WAVEOUT"
                    Case MM_ESS_ES1878_WAVEIN
                         sStr = "MM_ESS_ES1878_WAVEIN"
                    Case MM_ESS_ES1878_MIXER
                         sStr = "MM_ESS_ES1878_MIXER"
               End Select
          Case MM_AUDIOFILE ' Audio, Inc.
               sStr = "No Associated Product ID"
          Case MM_MOTOROLA ' Motorola, Inc.
               sStr = "No Associated Product ID"
          Case MM_CANOPUS ' Canopus, co., Ltd.
               sStr = "No Associated Product ID"
          Case MM_EPSON ' Seiko Epson Corporation
               sStr = "No Associated Product ID"
          Case MM_TRUEVISION ' Truevision
               Select Case lID
                    Case MM_TRUEVISION_WAVEIN1
                         sStr = "MM_TRUEVISION_WAVEIN1"
                    Case MM_TRUEVISION_WAVEOUT1
                         sStr = "MM_TRUEVISION_WAVEOUT1"
               End Select
          Case MM_AZTECH ' Aztech Labs, Inc.
               Select Case lID
                    Case MM_AZTECH_MIDIOUT
                         sStr = "MM_AZTECH_MIDIOUT"
                    Case MM_AZTECH_MIDIIN
                         sStr = "MM_AZTECH_MIDIIN"
                    Case MM_AZTECH_WAVEIN
                         sStr = "MM_AZTECH_WAVEIN"
                    Case MM_AZTECH_WAVEOUT
                         sStr = "MM_AZTECH_WAVEOUT"
                    Case MM_AZTECH_FMSYNTH
                         sStr = "MM_AZTECH_FMSYNTH"
                    Case MM_AZTECH_MIXER
                         sStr = "MM_AZTECH_MIXER"
                    Case MM_AZTECH_PRO16_WAVEIN
                         sStr = "MM_AZTECH_PRO16_WAVEIN"
                    Case MM_AZTECH_PRO16_WAVEOUT
                         sStr = "MM_AZTECH_PRO16_WAVEOUT"
                    Case MM_AZTECH_PRO16_FMSYNTH
                         sStr = "MM_AZTECH_PRO16_FMSYNTH"
                    Case MM_AZTECH_DSP16_WAVEIN
                         sStr = "MM_AZTECH_DSP16_WAVEIN"
                    Case MM_AZTECH_DSP16_WAVEOUT
                         sStr = "MM_AZTECH_DSP16_WAVEOUT"
                    Case MM_AZTECH_DSP16_FMSYNTH
                         sStr = "MM_AZTECH_DSP16_FMSYNTH"
                    Case MM_AZTECH_DSP16_WAVESYNTH
                         sStr = "MM_AZTECH_DSP16_WAVESYNTH"
                    Case MM_AZTECH_NOVA16_WAVEIN
                         sStr = "MM_AZTECH_NOVA16_WAVEIN"
                    Case MM_AZTECH_NOVA16_WAVEOUT
                         sStr = "MM_AZTECH_NOVA16_WAVEOUT"
                    Case MM_AZTECH_NOVA16_MIXER
                         sStr = "MM_AZTECH_NOVA16_MIXER"
                    Case MM_AZTECH_WASH16_WAVEIN
                         sStr = "MM_AZTECH_WASH16_WAVEIN"
                    Case MM_AZTECH_WASH16_WAVEOUT
                         sStr = "MM_AZTECH_WASH16_WAVEOUT"
                    Case MM_AZTECH_WASH16_MIXER
                          sStr = "MM_AZTECH_WASH16_MIXER"
                    Case MM_AZTECH_AUX_CD
                         sStr = "MM_AZTECH_AUX_CD"
                    Case MM_AZTECH_AUX_LINE
                         sStr = "MM_AZTECH_AUX_LINE"
                    Case MM_AZTECH_AUX_MIC
                         sStr = "MM_AZTECH_AUX_MIC"
                    Case MM_AZTECH_AUX
                         sStr = "MM_AZTECH_AUX"
               End Select
          Case MM_VIDEOLOGIC ' Videologic
               Select Case lID
                    Case MM_VIDEOLOGIC_MSWAVEIN
                         sStr = "MM_VIDEOLOGIC_MSWAVEIN"
                    Case MM_VIDEOLOGIC_MSWAVEOUT
                         sStr = "MM_VIDEOLOGIC_MSWAVEOUT"
               End Select
          Case MM_SCALACS ' SCALACS
               sStr = "No Associated Product ID"
          Case MM_KORG ' Korg Inc.
               Select Case lID
                    Case MM_KORG_PCIF_MIDIOUT
                         sStr = "MM_KORG_PCIF_MIDIOUT"
                    Case MM_KORG_PCIF_MIDIIN
                         sStr = "MM_KORG_PCIF_MIDIIN"
               End Select
          Case MM_APT ' Audio Processing Technology
               sStr = "MM_APT_ACE100CD"
          Case MM_ICS ' Integrated Circuit Systems, Inc.
               Select Case lID
                    Case MM_ICS_WAVEDECK_WAVEOUT
                         sStr = "MM_ICS_WAVEDECK_WAVEOUT"
                    Case MM_ICS_WAVEDECK_WAVEIN
                         sStr = "MM_ICS_WAVEDECK_WAVEIN"
                    Case MM_ICS_WAVEDECK_MIXER
                         sStr = "MM_ICS_WAVEDECK_MIXER"
                    Case MM_ICS_WAVEDECK_AUX
                         sStr = "MM_ICS_WAVEDECK_AUX"
                    Case MM_ICS_WAVEDECK_SYNTH
                         sStr = "MM_ICS_WAVEDECK_SYNTH"
                    Case MM_ICS_WAVEDEC_SB_WAVEOUT
                         sStr = "MM_ICS_WAVEDEC_SB_WAVEOUT"
                    Case MM_ICS_WAVEDEC_SB_WAVEIN
                         sStr = "MM_ICS_WAVEDEC_SB_WAVEIN"
                    Case MM_ICS_WAVEDEC_SB_FM_MIDIOUT
                         sStr = "MM_ICS_WAVEDEC_SB_FM_MIDIOUT"
                    Case MM_ICS_WAVEDEC_SB_MPU401_MIDIOUT
                         sStr = "MM_ICS_WAVEDEC_SB_MPU401_MIDIOUT"
                    Case MM_ICS_WAVEDEC_SB_MPU401_MIDIIN
                         sStr = "MM_ICS_WAVEDEC_SB_MPU401_MIDIIN"
                    Case MM_ICS_WAVEDEC_SB_MIXER
                         sStr = "MM_ICS_WAVEDEC_SB_MIXER"
                    Case MM_ICS_WAVEDEC_SB_AUX
                         sStr = "MM_ICS_WAVEDEC_SB_AUX"
                    Case MM_ICS_2115_LITE_MIDIOUT
                         sStr = "MM_ICS_2115_LITE_MIDIOUT"
                    Case MM_ICS_2120_LITE_MIDIOUT
                         sStr = "MM_ICS_2120_LITE_MIDIOUT"
               End Select
          Case MM_ITERATEDSYS ' Iterated Systems, Inc.
               sStr = "MM_ITERATEDSYS_FUFCODEC"
          Case MM_METHEUS ' Metheus
               sStr = "MM_METHEUS_ZIPPER"
          Case MM_LOGITECH ' Logitech, Inc.
               sStr = "No Associated Product ID"
          Case MM_WINNOV ' Winnov, Inc.
               Select Case lID
                    Case MM_WINNOV_CAVIAR_WAVEIN
                         sStr = "MM_WINNOV_CAVIAR_WAVEIN"
                    Case MM_WINNOV_CAVIAR_WAVEOUT
                         sStr = "MM_WINNOV_CAVIAR_WAVEOUT"
                    Case MM_WINNOV_CAVIAR_VIDC
                         sStr = "MM_WINNOV_CAVIAR_VIDC"
                    Case MM_WINNOV_CAVIAR_CHAMPAGNE
                         sStr = "MM_WINNOV_CAVIAR_CHAMPAGNE"
                    Case MM_WINNOV_CAVIAR_YUV8
                         sStr = "MM_WINNOV_CAVIAR_YUV8"
               End Select
          Case MM_NCR ' NCR Corporation
               Select Case lID
                    Case MM_NCR_BA_WAVEIN
                         sStr = "MM_NCR_BA_WAVEIN"
                    Case MM_NCR_BA_WAVEOUT
                         sStr = "MM_NCR_BA_WAVEOUT"
                    Case MM_NCR_BA_SYNTH
                         sStr = "MM_NCR_BA_SYNTH"
                    Case MM_NCR_BA_AUX
                         sStr = "MM_NCR_BA_AUX"
                    Case MM_NCR_BA_MIXER
                         sStr = "MM_NCR_BA_MIXER"
               End Select
          Case MM_EXAN ' EXAN
               sStr = "No Associated Product ID"
          Case MM_AST ' AST Research Inc.
               Select Case lID
                    Case MM_AST_MODEMWAVE_WAVEIN
                         sStr = "MM_AST_MODEMWAVE_WAVEIN"
                    Case MM_AST_MODEMWAVE_WAVEOUT
                         sStr = "MM_AST_MODEMWAVE_WAVEOUT"
               End Select
          Case MM_WILLOWPOND ' Willow Pond Corporation
               sStr = "No Associated Product ID"
          Case MM_SONICFOUNDRY ' Sonic Foundry
               sStr = "No Associated Product ID"
          Case MM_VITEC ' Vitec Multimedia
               Select Case lID
                    Case MM_VITEC_VMAKER
                         sStr = "MM_VITEC_VMAKER"
                    Case MM_VITEC_VMPRO
                         sStr = "MM_VITEC_VMPRO"
               End Select
          Case MM_MOSCOM ' MOSCOM Corporation
               Select Case lID
                    Case MM_MOSCOM_VPC2400_IN
                         sStr = "MM_MOSCOM_VPC2400_IN"
                    Case MM_MOSCOM_VPC2400_OUT
                         sStr = "MM_MOSCOM_VPC2400_OUT"
               End Select
          Case MM_SILICONSOFT ' Silicon Soft, Inc.
               Select Case lID
                    Case MM_SILICONSOFT_SC1_WAVEIN
                         sStr = "MM_SILICONSOFT_SC1_WAVEIN"
                    Case MM_SILICONSOFT_SC1_WAVEOUT
                         sStr = "MM_SILICONSOFT_SC1_WAVEOUT"
                    Case MM_SILICONSOFT_SC2_WAVEIN
                         sStr = "MM_SILICONSOFT_SC2_WAVEIN"
                    Case MM_SILICONSOFT_SC2_WAVEOUT
                         sStr = "MM_SILICONSOFT_SC2_WAVEOUT"
                    Case MM_SILICONSOFT_SOUNDJR2_WAVEOUT
                         sStr = "MM_SILICONSOFT_SOUNDJR2_WAVEOUT"
                    Case MM_SILICONSOFT_SOUNDJR2PR_WAVEIN
                         sStr = "MM_SILICONSOFT_SOUNDJR2PR_WAVEIN"
                    Case MM_SILICONSOFT_SOUNDJR2PR_WAVEOUT
                         sStr = "MM_SILICONSOFT_SOUNDJR2PR_WAVEOUT"
                    Case MM_SILICONSOFT_SOUNDJR3_WAVEOUT
                         sStr = "MM_SILICONSOFT_SOUNDJR3_WAVEOUT"
               End Select
          Case MM_SUPERMAC ' Supermac
               sStr = "No Associated Product ID"
          Case MM_AUDIOPT ' Audio Processing Technology
               sStr = "No Associated Product ID"
          Case MM_SPEECHCOMP ' Speech Compression
               sStr = "No Associated Product ID"
          Case MM_AHEAD ' Ahead, Inc.
               Select Case lID
                    Case MM_AHEAD_MULTISOUND
                         sStr = "MM_AHEAD_MULTISOUND"
                    Case MM_AHEAD_SOUNDBLASTER
                         sStr = "MM_AHEAD_SOUNDBLASTER"
                    Case MM_AHEAD_PROAUDIO
                         sStr = "MM_AHEAD_PROAUDIO"
                    Case MM_AHEAD_GENERIC
                         sStr = "MM_AHEAD_GENERIC"
               End Select
          Case MM_DOLBY ' Dolby Laboratories
               sStr = "No Associated Product ID"
          Case MM_OKI ' OKI
               sStr = "No Associated Product ID"
          Case MM_AURAVISION ' AuraVision Corporation
               sStr = "No Associated Product ID"
          Case MM_OLIVETTI ' Ing C. Olivetti & C., S.p.A.
               Select Case lID
                    Case MM_OLIVETTI_WAVEIN
                         sStr = "M_OLIVETTI_WAVEIN"
                    Case MM_OLIVETTI_WAVEOUT
                         sStr = "MM_OLIVETTI_WAVEOUT"
                    Case MM_OLIVETTI_MIXER
                         sStr = "MM_OLIVETTI_MIXER"
                    Case MM_OLIVETTI_AUX
                         sStr = "MM_OLIVETTI_AUX"
                    Case MM_OLIVETTI_MIDIIN
                         sStr = "MM_OLIVETTI_MIDIIN"
                    Case MM_OLIVETTI_MIDIOUT
                         sStr = "MM_OLIVETTI_MIDIOUT"
                    Case MM_OLIVETTI_SYNTH
                         sStr = "MM_OLIVETTI_SYNTH"
                    Case MM_OLIVETTI_JOYSTICK
                         sStr = "MM_OLIVETTI_JOYSTICK"
                    Case MM_OLIVETTI_ACM_GSM
                         sStr = "MM_OLIVETTI_ACM_GSM"
                    Case MM_OLIVETTI_ACM_ADPCM
                         sStr = "MM_OLIVETTI_ACM_ADPCM"
                    Case MM_OLIVETTI_ACM_CELP
                         sStr = "MM_OLIVETTI_ACM_CELP"
                    Case MM_OLIVETTI_ACM_SBC
                         sStr = "MM_OLIVETTI_ACM_SBC"
                    Case MM_OLIVETTI_ACM_OPR
                         sStr = "MM_OLIVETTI_ACM_OPR"
               End Select
          Case MM_IOMAGIC ' I/O Magic Corporation
               Select Case lID
                    Case MM_IOMAGIC_TEMPO_WAVEOUT
                         sStr = "MM_IOMAGIC_TEMPO_WAVEOUT"
                    Case MM_IOMAGIC_TEMPO_WAVEIN
                         sStr = "MM_IOMAGIC_TEMPO_WAVEIN"
                    Case MM_IOMAGIC_TEMPO_SYNTH
                         sStr = "MM_IOMAGIC_TEMPO_SYNTH"
                    Case MM_IOMAGIC_TEMPO_MIDIOUT
                         sStr = "MM_IOMAGIC_TEMPO_MIDIOUT"
                    Case MM_IOMAGIC_TEMPO_MXDOUT
                         sStr = "MM_IOMAGIC_TEMPO_MXDOUT"
                    Case MM_IOMAGIC_TEMPO_AUXOUT
                         sStr = "MM_IOMAGIC_TEMPO_AUXOUT"
               End Select
          Case MM_MATSUSHITA ' Matsushita Electric Industrial Co., LTD.
               Select Case lID
                    Case MM_MATSUSHITA_WAVEIN
                         sStr = "MM_MATSUSHITA_WAVEIN"
                    Case MM_MATSUSHITA_WAVEOUT
                         sStr = "MM_MATSUSHITA_WAVEOUT"
                    Case MM_MATSUSHITA_FMSYNTH_STEREO
                         sStr = "MM_MATSUSHITA_FMSYNTH_STEREO"
                    Case MM_MATSUSHITA_MIXER
                          sStr = "MM_MATSUSHITA_MIXER"
                   Case MM_MATSUSHITA_AUX
                         sStr = "MM_MATSUSHITA_AUX"
               End Select
          Case MM_CONTROLRES ' Control Resources Limited
               sStr = "No Associated Product ID"
          Case MM_XEBEC ' Xebec Multimedia Solutions Limited
               sStr = "No Associated Product ID"
          Case MM_NEWMEDIA ' New Media Corporation
               sStr = "MM_NEWMEDIA_WAVJAMMER"
          Case MM_NMS ' Natural MicroSystems
               sStr = "No Associated Product ID"
          Case MM_LYRRUS ' Lyrrus Inc.
               sStr = "MM_LYRRUS_BRIDGE_GUITAR"
          Case MM_COMPUSIC ' Compusic
               sStr = "No Associated Product ID"
          Case MM_OPTI ' OPTi Computers Inc.
               Select Case lID
                    Case MM_OPTI_M16_FMSYNTH_STEREO
                         sStr = "MM_OPTI_M16_FMSYNTH_STEREO"
                    Case MM_OPTI_M16_MIDIIN
                         sStr = "MM_OPTI_M16_MIDIIN"
                    Case MM_OPTI_M16_MIDIOUT
                         sStr = "MM_OPTI_M16_MIDIOUT"
                    Case MM_OPTI_M16_WAVEIN
                         sStr = "MM_OPTI_M16_WAVEIN"
                    Case MM_OPTI_M16_WAVEOUT
                         sStr = "MM_OPTI_M16_WAVEOUT"
                    Case MM_OPTI_M16_MIXER
                         sStr = "MM_OPTI_M16_MIXER"
                    Case MM_OPTI_M16_AUX
                         sStr = "MM_OPTI_M16_AUX"
                    Case MM_OPTI_P16_FMSYNTH_STEREO
                         sStr = "MM_OPTI_P16_FMSYNTH_STEREO"
                    Case MM_OPTI_P16_MIDIIN
                         sStr = "MM_OPTI_P16_MIDIIN"
                    Case MM_OPTI_P16_MIDIOUT
                         sStr = "MM_OPTI_P16_MIDIOUT"
                    Case MM_OPTI_P16_WAVEIN
                         sStr = "MM_OPTI_P16_WAVEIN"
                    Case MM_OPTI_P16_WAVEOUT
                         sStr = "MM_OPTI_P16_WAVEOUT"
                    Case MM_OPTI_P16_MIXER
                         sStr = "MM_OPTI_P16_MIXER"
                    Case MM_OPTI_P16_AUX
                         sStr = "MM_OPTI_P16_AUX"
                    Case MM_OPTI_M32_WAVEIN
                         sStr = "MM_OPTI_M32_WAVEIN"
                    Case MM_OPTI_M32_WAVEOUT
                         sStr = "MM_OPTI_M32_WAVEOUT"
                    Case MM_OPTI_M32_MIDIIN
                         sStr = "MM_OPTI_M32_MIDIIN"
                    Case MM_OPTI_M32_MIDIOUT
                         sStr = "MM_OPTI_M32_MIDIOUT"
                    Case MM_OPTI_M32_SYNTH_STEREO
                         sStr = "MM_OPTI_M32_SYNTH_STEREO"
                    Case MM_OPTI_M32_MIXER
                         sStr = "MM_OPTI_M32_MIXER"
                    Case MM_OPTI_M32_AUX
                         sStr = "MM_OPTI_M32_AUX"
               End Select
          Case MM_ADLACC ' Adlib Accessories Inc.
               sStr = "No Associated Product ID"
          Case MM_COMPAQ ' Compaq Computer Corp.
               Select Case lID
                    Case MM_COMPAQ_BB_WAVEIN
                         sStr = "MM_COMPAQ_BB_WAVEIN"
                    Case MM_COMPAQ_BB_WAVEOUT
                         sStr = "MM_COMPAQ_BB_WAVEOUT"
                    Case MM_COMPAQ_BB_WAVEAUX
                         sStr = "MM_COMPAQ_BB_WAVEAUX"
               End Select
          Case MM_DIALOGIC ' Dialogic Corporation
               sStr = "No Associated Product ID"
          Case MM_INSOFT ' InSoft, Inc.
               sStr = "No Associated Product ID"
          Case MM_MPTUS ' M.P. Technologies, Inc.
               sStr = "MM_MPTUS_SPWAVEOUT"
          Case MM_WEITEK ' Weitek
               sStr = "No Associated Product ID"
          Case MM_LERNOUT_AND_HAUSPIE ' Lernout & Hauspie
               sStr = "MM_LERNOUT_ANDHAUSPIE_LHCODECACM"
          Case MM_QCIAR ' Quanta Computer Inc.
               sStr = "No Associated Product ID"
          Case MM_APPLE ' Apple Computer, Inc.
               sStr = "No Associated Product ID"
          Case MM_DIGITAL ' Digital Equipment Corporation
               Select Case lID
                    Case MM_DIGITAL_AV320_WAVEIN
                         sStr = "MM_DIGITAL_AV320_WAVEIN"
                    Case MM_DIGITAL_AV320_WAVEOUT
                         sStr = "MM_DIGITAL_AV320_WAVEOUT"
               End Select
          Case MM_MOTU ' Mark of the Unicorn
               Select Case lID
                    Case MM_MOTU_MTP_MIDIOUT_ALL
                         sStr = "MM_MOTU_MTP_MIDIOUT_ALL"
                    Case MM_MOTU_MTP_MIDIIN_1
                         sStr = "MM_MOTU_MTP_MIDIIN_1"
                    Case MM_MOTU_MTP_MIDIOUT_1
                         sStr = "MM_MOTU_MTP_MIDIOUT_1"
                    Case MM_MOTU_MTP_MIDIIN_2
                         sStr = "MM_MOTU_MTP_MIDIIN_2"
                    Case MM_MOTU_MTP_MIDIOUT_2
                         sStr = "MM_MOTU_MTP_MIDIOUT_2"
                    Case MM_MOTU_MTP_MIDIIN_3
                         sStr = "MM_MOTU_MTP_MIDIIN_3"
                    Case MM_MOTU_MTP_MIDIOUT_3
                         sStr = "MM_MOTU_MTP_MIDIOUT_3"
                    Case MM_MOTU_MTP_MIDIIN_4
                         sStr = "MM_MOTU_MTP_MIDIIN_4"
                    Case MM_MOTU_MTP_MIDIOUT_4
                         sStr = "MM_MOTU_MTP_MIDIOUT_4"
                    Case MM_MOTU_MTP_MIDIIN_5
                         sStr = "MM_MOTU_MTP_MIDIIN_5"
                    Case MM_MOTU_MTP_MIDIOUT_5
                         sStr = "MM_MOTU_MTP_MIDIOUT_5"
                    Case MM_MOTU_MTP_MIDIIN_6
                         sStr = "MM_MOTU_MTP_MIDIIN_6"
                    Case MM_MOTU_MTP_MIDIOUT_6
                         sStr = "MM_MOTU_MTP_MIDIOUT_6"
                    Case MM_MOTU_MTP_MIDIIN_7
                         sStr = "MM_MOTU_MTP_MIDIIN_7"
                    Case MM_MOTU_MTP_MIDIOUT_7
                         sStr = "MM_MOTU_MTP_MIDIOUT_7"
                    Case MM_MOTU_MTP_MIDIIN_8
                         sStr = "MM_MOTU_MTP_MIDIIN_8"
                    Case MM_MOTU_MTP_MIDIOUT_8
                         sStr = "MM_MOTU_MTP_MIDIOUT_8"
                    Case MM_MOTU_MTPII_MIDIOUT_ALL
                         sStr = "MM_MOTU_MTPII_MIDIOUT_ALL"
                    Case MM_MOTU_MTPII_MIDIIN_SYNC
                         sStr = "MM_MOTU_MTPII_MIDIIN_SYNC"
                    Case MM_MOTU_MTPII_MIDIIN_1
                         sStr = "MM_MOTU_MTPII_MIDIIN_1"
                    Case MM_MOTU_MTPII_MIDIOUT_1
                         sStr = "MM_MOTU_MTPII_MIDIOUT_1"
                    Case MM_MOTU_MTPII_MIDIIN_2
                         sStr = "MM_MOTU_MTPII_MIDIIN_2"
                    Case MM_MOTU_MTPII_MIDIOUT_2
                         sStr = "MM_MOTU_MTPII_MIDIOUT_2"
                    Case MM_MOTU_MTPII_MIDIIN_3
                         sStr = "MM_MOTU_MTPII_MIDIIN_3"
                    Case MM_MOTU_MTPII_MIDIOUT_3
                         sStr = "MM_MOTU_MTPII_MIDIOUT_3"
                    Case MM_MOTU_MTPII_MIDIIN_4
                         sStr = "M_MOTU_MTPII_MIDIIN_4"
                    Case MM_MOTU_MTPII_MIDIOUT_4
                         sStr = "M_MOTU_MTPII_MIDIOUT_4"
                    Case MM_MOTU_MTPII_MIDIIN_5
                         sStr = "MM_MOTU_MTPII_MIDIIN_5"
                    Case MM_MOTU_MTPII_MIDIOUT_5
                         sStr = "MM_MOTU_MTPII_MIDIOUT_5"
                    Case MM_MOTU_MTPII_MIDIIN_6
                         sStr = "MM_MOTU_MTPII_MIDIIN_6"
                    Case MM_MOTU_MTPII_MIDIOUT_6
                         sStr = "MM_MOTU_MTPII_MIDIOUT_6"
                    Case MM_MOTU_MTPII_MIDIIN_7
                         sStr = "MM_MOTU_MTPII_MIDIIN_7"
                    Case MM_MOTU_MTPII_MIDIOUT_7
                         sStr = "MM_MOTU_MTPII_MIDIOUT_7"
                    Case MM_MOTU_MTPII_MIDIIN_8
                         sStr = "MM_MOTU_MTPII_MIDIIN_8"
                    Case MM_MOTU_MTPII_MIDIOUT_8
                         sStr = "MM_MOTU_MTPII_MIDIOUT_8"
                    Case MM_MOTU_MTPII_NET_MIDIIN_1
                         sStr = "MM_MOTU_MTPII_NET_MIDIIN_1"
                    Case MM_MOTU_MTPII_NET_MIDIOUT_1
                         sStr = "MM_MOTU_MTPII_NET_MIDIOUT_1"
                    Case MM_MOTU_MTPII_NET_MIDIIN_2
                         sStr = "MM_MOTU_MTPII_NET_MIDIIN_2"
                    Case MM_MOTU_MTPII_NET_MIDIOUT_2
                         sStr = "MM_MOTU_MTPII_NET_MIDIOUT_2"
                    Case MM_MOTU_MTPII_NET_MIDIIN_3
                         sStr = "MM_MOTU_MTPII_NET_MIDIIN_3"
                    Case MM_MOTU_MTPII_NET_MIDIOUT_3
                         sStr = " MM_MOTU_MTPII_NET_MIDIOUT_3"
                    Case MM_MOTU_MTPII_NET_MIDIIN_4
                         sStr = "MM_MOTU_MTPII_NET_MIDIIN_4"
                    Case MM_MOTU_MTPII_NET_MIDIOUT_4
                         sStr = "MM_MOTU_MTPII_NET_MIDIOUT_4"
                    Case MM_MOTU_MTPII_NET_MIDIIN_5
                         sStr = "MM_MOTU_MTPII_NET_MIDIIN_5"
                    Case MM_MOTU_MTPII_NET_MIDIOUT_5
                         sStr = "MM_MOTU_MTPII_NET_MIDIOUT_5"
                    Case MM_MOTU_MTPII_NET_MIDIIN_6
                         sStr = "MM_MOTU_MTPII_NET_MIDIIN_6"
                    Case MM_MOTU_MTPII_NET_MIDIOUT_6
                         sStr = "MM_MOTU_MTPII_NET_MIDIOUT_6"
                    Case MM_MOTU_MTPII_NET_MIDIIN_7
                         sStr = "MM_MOTU_MTPII_NET_MIDIIN_7"
                    Case MM_MOTU_MTPII_NET_MIDIOUT_7
                         sStr = "MM_MOTU_MTPII_NET_MIDIOUT_7"
                    Case MM_MOTU_MTPII_NET_MIDIIN_8
                         sStr = "MM_MOTU_MTPII_NET_MIDIIN_8"
                    Case MM_MOTU_MTPII_NET_MIDIOUT_8
                         sStr = "MM_MOTU_MTPII_NET_MIDIOUT_8"
                    Case MM_MOTU_MXP_MIDIIN_MIDIOUT_ALL
                         sStr = "MM_MOTU_MXP_MIDIIN_MIDIOUT_ALL"
                    Case MM_MOTU_MXP_MIDIIN_SYNC
                         sStr = "MM_MOTU_MXP_MIDIIN_SYNC"
                    Case MM_MOTU_MXP_MIDIIN_MIDIIN_1
                         sStr = "MM_MOTU_MXP_MIDIIN_MIDIIN_1"
                    Case MM_MOTU_MXP_MIDIIN_MIDIOUT_1
                         sStr = "MM_MOTU_MXP_MIDIIN_MIDIOUT_1"
                    Case MM_MOTU_MXP_MIDIIN_MIDIIN_2
                         sStr = "MM_MOTU_MXP_MIDIIN_MIDIIN_2"
                    Case MM_MOTU_MXP_MIDIIN_MIDIOUT_2 = 302
                         sStr = ""
                    Case MM_MOTU_MXP_MIDIIN_MIDIIN_3 = 303
                         sStr = "MM_MOTU_MXP_MIDIIN_MIDIIN_3"
                    Case MM_MOTU_MXP_MIDIIN_MIDIOUT_3 = 303
                         sStr = "MM_MOTU_MXP_MIDIIN_MIDIOUT_3"
                    Case MM_MOTU_MXP_MIDIIN_MIDIIN_4 = 304
                         sStr = "MM_MOTU_MXP_MIDIIN_MIDIIN_4"
                    Case MM_MOTU_MXP_MIDIIN_MIDIOUT_4 = 304
                         sStr = "MM_MOTU_MXP_MIDIIN_MIDIOUT_4"
                    Case MM_MOTU_MXP_MIDIIN_MIDIIN_5 = 305
                         sStr = "MM_MOTU_MXP_MIDIIN_MIDIIN_5"
                    Case MM_MOTU_MXP_MIDIIN_MIDIOUT_5 = 305
                         sStr = "MM_MOTU_MXP_MIDIIN_MIDIOUT_5"
                    Case MM_MOTU_MXP_MIDIIN_MIDIIN_6 = 306
                         sStr = "MM_MOTU_MXP_MIDIIN_MIDIIN_6"
                    Case MM_MOTU_MXP_MIDIIN_MIDIOUT_6 = 306
                         sStr = "MM_MOTU_MXP_MIDIIN_MIDIOUT_6"
                    Case MM_MOTU_MXPMPU_MIDIOUT_ALL = 400
                         sStr = "MM_MOTU_MXPMPU_MIDIOUT_ALL"
                    Case MM_MOTU_MXPMPU_MIDIIN_SYNC = 400
                         sStr = "M_MOTU_MXPMPU_MIDIIN_SYNC"
                    Case MM_MOTU_MXPMPU_MIDIIN_1 = 401
                         sStr = "MM_MOTU_MXPMPU_MIDIIN_1"
                    Case MM_MOTU_MXPMPU_MIDIOUT_1 = 401
                         sStr = "MM_MOTU_MXPMPU_MIDIOUT_1"
                    Case MM_MOTU_MXPMPU_MIDIIN_2 = 402
                         sStr = "MM_MOTU_MXPMPU_MIDIIN_2"
                    Case MM_MOTU_MXPMPU_MIDIOUT_2 = 402
                         sStr = "MM_MOTU_MXPMPU_MIDIOUT_2"
                    Case MM_MOTU_MXPMPU_MIDIIN_3 = 403
                         sStr = "MM_MOTU_MXPMPU_MIDIIN_3"
                    Case MM_MOTU_MXPMPU_MIDIOUT_3 = 403
                         sStr = "MM_MOTU_MXPMPU_MIDIOUT_3"
                    Case MM_MOTU_MXPMPU_MIDIIN_4 = 404
                         sStr = "MM_MOTU_MXPMPU_MIDIIN_4"
                    Case MM_MOTU_MXPMPU_MIDIOUT_4 = 404
                         sStr = "MM_MOTU_MXPMPU_MIDIOUT_4"
                    Case MM_MOTU_MXPMPU_MIDIIN_5 = 405
                         sStr = "MM_MOTU_MXPMPU_MIDIIN_5"
                    Case MM_MOTU_MXPMPU_MIDIOUT_5 = 405
                         sStr = "MM_MOTU_MXPMPU_MIDIOUT_5"
                    Case MM_MOTU_MXPMPU_MIDIIN_6 = 406
                         sStr = "MM_MOTU_MXPMPU_MIDIIN_6"
                    Case MM_MOTU_MXPMPU_MIDIOUT_6 = 406
                         sStr = "MM_MOTU_MXPMPU_MIDIOUT_6"
                    Case MM_MOTU_MXN_MIDIOUT_ALL = 500
                         sStr = "MM_MOTU_MXN_MIDIOUT_ALL"
                    Case MM_MOTU_MXN_MIDIIN_SYNC = 500
                         sStr = "MM_MOTU_MXN_MIDIIN_SYNC"
                    Case MM_MOTU_MXN_MIDIIN_1 = 501
                         sStr = "MM_MOTU_MXN_MIDIIN_1"
                    Case MM_MOTU_MXN_MIDIOUT_1 = 501
                         sStr = "MM_MOTU_MXN_MIDIOUT_1"
                    Case MM_MOTU_MXN_MIDIIN_2 = 502
                         sStr = "MM_MOTU_MXN_MIDIIN_2"
                    Case MM_MOTU_MXN_MIDIOUT_2 = 502
                         sStr = "MM_MOTU_MXN_MIDIOUT_2"
                    Case MM_MOTU_MXN_MIDIIN_3 = 503
                         sStr = "MM_MOTU_MXN_MIDIIN_3"
                    Case MM_MOTU_MXN_MIDIOUT_3 = 503
                         sStr = "MM_MOTU_MXN_MIDIOUT_3"
                    Case MM_MOTU_MXN_MIDIIN_4 = 504
                         sStr = "MM_MOTU_MXN_MIDIIN_4"
                    Case MM_MOTU_MXN_MIDIOUT_4 = 504
                         sStr = "MM_MOTU_MXN_MIDIOUT_4"
                    Case MM_MOTU_FLYER_MIDI_IN_SYNC = 600
                         sStr = "MM_MOTU_FLYER_MIDI_IN_SYNC"
                    Case MM_MOTU_FLYER_MIDI_IN_A = 601
                         sStr = "MM_MOTU_FLYER_MIDI_IN_A"
                    Case MM_MOTU_FLYER_MIDI_OUT_A = 601
                         sStr = "MM_MOTU_FLYER_MIDI_OUT_A"
                    Case MM_MOTU_FLYER_MIDI_IN_B = 602
                         sStr = "M_MOTU_FLYER_MIDI_IN_B"
                    Case MM_MOTU_FLYER_MIDI_OUT_B = 602
                         sStr = "MM_MOTU_FLYER_MIDI_OUT_B"
                    Case MM_MOTU_PKX_MIDI_IN_SYNC = 700
                         sStr = "MM_MOTU_PKX_MIDI_IN_SYNC"
                    Case MM_MOTU_PKX_MIDI_IN_A = 701
                         sStr = "MM_MOTU_PKX_MIDI_IN_A"
                    Case MM_MOTU_PKX_MIDI_OUT_A = 701
                         sStr = "MM_MOTU_PKX_MIDI_OUT_A"
                    Case MM_MOTU_PKX_MIDI_IN_B = 702
                         sStr = "MM_MOTU_PKX_MIDI_IN_B"
                    Case MM_MOTU_PKX_MIDI_OUT_B = 702
                         sStr = "MM_MOTU_PKX_MIDI_OUT_B"
                    Case MM_MOTU_DTX_MIDI_IN_SYNC = 800
                         sStr = "MM_MOTU_DTX_MIDI_IN_SYNC"
                    Case MM_MOTU_DTX_MIDI_IN_A = 801
                         sStr = "MM_MOTU_DTX_MIDI_IN_A"
                    Case MM_MOTU_DTX_MIDI_OUT_A = 801
                         sStr = "MM_MOTU_DTX_MIDI_OUT_A"
                    Case MM_MOTU_DTX_MIDI_IN_B = 802
                         sStr = "MM_MOTU_DTX_MIDI_IN_B"
                    Case MM_MOTU_DTX_MIDI_OUT_B = 802
                         sStr = "MM_MOTU_DTX_MIDI_OUT_B"
                    Case MM_MOTU_MTPAV_MIDIOUT_ALL = 900
                         sStr = "MM_MOTU_MTPAV_MIDIOUT_ALL"
                    Case MM_MOTU_MTPAV_MIDIIN_SYNC = 900
                         sStr = "MM_MOTU_MTPAV_MIDIIN_SYNC"
                    Case MM_MOTU_MTPAV_MIDIIN_1 = 901
                         sStr = "MM_MOTU_MTPAV_MIDIIN_1"
                    Case MM_MOTU_MTPAV_MIDIOUT_1 = 901
                         sStr = "MM_MOTU_MTPAV_MIDIOUT_1"
                    Case MM_MOTU_MTPAV_MIDIIN_2 = 902
                         sStr = "MM_MOTU_MTPAV_MIDIIN_2"
                    Case MM_MOTU_MTPAV_MIDIOUT_2 = 902
                         sStr = "MM_MOTU_MTPAV_MIDIOUT_2"
                    Case MM_MOTU_MTPAV_MIDIIN_3 = 903
                         sStr = "MM_MOTU_MTPAV_MIDIIN_3"
                    Case MM_MOTU_MTPAV_MIDIOUT_3 = 903
                         sStr = "MM_MOTU_MTPAV_MIDIOUT_3"
                    Case MM_MOTU_MTPAV_MIDIIN_4 = 904
                         sStr = "MM_MOTU_MTPAV_MIDIIN_4"
                    Case MM_MOTU_MTPAV_MIDIOUT_4 = 904
                         sStr = "MM_MOTU_MTPAV_MIDIOUT_4"
                    Case MM_MOTU_MTPAV_MIDIIN_5 = 905
                         sStr = "MM_MOTU_MTPAV_MIDIIN_5"
                    Case MM_MOTU_MTPAV_MIDIOUT_5 = 905
                         sStr = "MM_MOTU_MTPAV_MIDIOUT_5"
                    Case MM_MOTU_MTPAV_MIDIIN_6 = 906
                         sStr = "MM_MOTU_MTPAV_MIDIIN_6"
                    Case MM_MOTU_MTPAV_MIDIOUT_6 = 906
                         sStr = "MM_MOTU_MTPAV_MIDIOUT_6"
                    Case MM_MOTU_MTPAV_MIDIIN_7 = 907
                         sStr = "MM_MOTU_MTPAV_MIDIIN_7"
                    Case MM_MOTU_MTPAV_MIDIOUT_7 = 907
                         sStr = "MM_MOTU_MTPAV_MIDIOUT_7"
                    Case MM_MOTU_MTPAV_MIDIIN_8 = 908
                         sStr = "MM_MOTU_MTPAV_MIDIIN_8"
                    Case MM_MOTU_MTPAV_MIDIOUT_8 = 908
                         sStr = "MM_MOTU_MTPAV_MIDIOUT_8"
                    Case MM_MOTU_MTPAV_NET_MIDIIN_1 = 909
                         sStr = "MM_MOTU_MTPAV_NET_MIDIIN_1"
                    Case MM_MOTU_MTPAV_NET_MIDIOUT_1 = 909
                         sStr = "MM_MOTU_MTPAV_NET_MIDIOUT_1"
                    Case MM_MOTU_MTPAV_NET_MIDIIN_2 = 910
                         sStr = "MM_MOTU_MTPAV_NET_MIDIIN_2"
                    Case MM_MOTU_MTPAV_NET_MIDIOUT_2 = 910
                         sStr = "MM_MOTU_MTPAV_NET_MIDIOUT_2"
                    Case MM_MOTU_MTPAV_NET_MIDIIN_3 = 911
                         sStr = "MM_MOTU_MTPAV_NET_MIDIIN_3"
                    Case MM_MOTU_MTPAV_NET_MIDIOUT_3 = 911
                         sStr = "MM_MOTU_MTPAV_NET_MIDIOUT_3"
                    Case MM_MOTU_MTPAV_NET_MIDIIN_4 = 912
                         sStr = "MM_MOTU_MTPAV_NET_MIDIIN_4"
                    Case MM_MOTU_MTPAV_NET_MIDIOUT_4 = 912
                         sStr = "MM_MOTU_MTPAV_NET_MIDIOUT_4"
                    Case MM_MOTU_MTPAV_NET_MIDIIN_5 = 913
                         sStr = "MM_MOTU_MTPAV_NET_MIDIIN_5"
                    Case MM_MOTU_MTPAV_NET_MIDIOUT_5 = 913
                         sStr = "MM_MOTU_MTPAV_NET_MIDIOUT_5"
                    Case MM_MOTU_MTPAV_NET_MIDIIN_6 = 914
                         sStr = "MM_MOTU_MTPAV_NET_MIDIIN_6"
                    Case MM_MOTU_MTPAV_NET_MIDIOUT_6 = 914
                         sStr = "MM_MOTU_MTPAV_NET_MIDIOUT_6"
                    Case MM_MOTU_MTPAV_NET_MIDIIN_7 = 915
                         sStr = "MM_MOTU_MTPAV_NET_MIDIIN_7"
                    Case MM_MOTU_MTPAV_NET_MIDIOUT_7 = 915
                         sStr = "MM_MOTU_MTPAV_NET_MIDIOUT_7"
                    Case MM_MOTU_MTPAV_NET_MIDIIN_8 = 916
                         sStr = "MM_MOTU_MTPAV_NET_MIDIIN_8"
                    Case MM_MOTU_MTPAV_NET_MIDIOUT_8 = 916
                         sStr = "MM_MOTU_MTPAV_NET_MIDIOUT_8"
                    Case MM_MOTU_MTPAV_MIDIIN_ADAT = 917
                         sStr = "MM_MOTU_MTPAV_MIDIIN_ADAT"
                    Case MM_MOTU_MTPAV_MIDIOUT_ADAT = 917
                         sStr = "MM_MOTU_MTPAV_MIDIOUT_ADAT"
               End Select
          Case MM_WORKBIT ' Workbit Corporation
               Select Case lID
                    Case MM_WORKBIT_MIXER
                         sStr = "MM_WORKBIT_MIXER"
                    Case MM_WORKBIT_WAVEOUT
                         sStr = "MM_WORKBIT_WAVEOUT"
                    Case MM_WORKBIT_WAVEIN
                         sStr = "MM_WORKBIT_WAVEIN"
                    Case MM_WORKBIT_MIDIIN
                         sStr = "MM_WORKBIT_MIDIIN"
                    Case MM_WORKBIT_MIDIOUT
                         sStr = "MM_WORKBIT_MIDIOUT"
                    Case MM_WORKBIT_FMSYNTH
                         sStr = "MM_WORKBIT_FMSYNTH"
                    Case MM_WORKBIT_AUX
                         sStr = "MM_WORKBIT_AUX"
                    Case MM_WORKBIT_JOYSTICK
                         sStr = "MM_WORKBIT_JOYSTICK"
               End Select
          Case MM_OSITECH ' Ositech Communications Inc.
               sStr = "MM_OSITECH_TRUMPCARD"
          Case MM_MIRO ' miro Computer Products AG
               Select Case lID
                    Case MM_MIRO_MOVIEPRO
                         sStr = "MM_MIRO_MOVIEPRO"
                    Case MM_MIRO_VIDEOD1
                         sStr = "MM_MIRO_VIDEOD1"
                    Case MM_MIRO_VIDEODC1TV
                         sStr = "MM_MIRO_VIDEODC1TV"
                    Case MM_MIRO_VIDEOTD
                         sStr = "MM_MIRO_VIDEOTD"
                    Case MM_MIRO_DC30_WAVEOUT
                         sStr = "MM_MIRO_DC30_WAVEOUT"
                    Case MM_MIRO_DC30_WAVEIN
                         sStr = "MM_MIRO_DC30_WAVEIN"
                    Case MM_MIRO_DC30_MIX
                         sStr = "MM_MIRO_DC30_MIX"
               End Select
          Case MM_CIRRUSLOGIC ' Cirrus Logic
               sStr = "No Associated Product ID"
          Case MM_ISOLUTION ' ISOLUTION  B.V.
               sStr = "No Associated Product ID"
          Case MM_HORIZONS ' Horizons Technology, Inc
               sStr = "No Associated Product ID"
          Case MM_CONCEPTS ' Computer Concepts Ltd
               sStr = "No Associated Product ID"
          Case MM_VTG ' Voice Technologies Group, Inc.
               sStr = "No Associated Product ID"
          Case MM_RADIUS ' Radius
               sStr = "No Associated Product ID"
          Case MM_ROCKWELL ' Rockwell International
               Select Case lID
                    Case MM_VOICEMIXER
                         sStr = "MM_VOICEMIXER"
                    Case ROCKWELL_WA1_WAVEIN
                         sStr = "ROCKWELL_WA1_WAVEIN"
                    Case ROCKWELL_WA1_WAVEOUT
                         sStr = "ROCKWELL_WA1_WAVEOUT"
                    Case ROCKWELL_WA1_SYNTH
                         sStr = "ROCKWELL_WA1_SYNTH"
                    Case ROCKWELL_WA1_MIXER
                         sStr = "ROCKWELL_WA1_MIXER"
                    Case ROCKWELL_WA1_MPU401_IN
                         sStr = "ROCKWELL_WA1_MPU401_IN"
                    Case ROCKWELL_WA1_MPU401_OUT
                         sStr = "ROCKWELL_WA1_MPU401_OUT"
                    Case ROCKWELL_WA2_WAVEIN
                         sStr = "ROCKWELL_WA2_WAVEIN"
                    Case ROCKWELL_WA2_WAVEOUT
                         sStr = "ROCKWELL_WA2_WAVEOUT"
                    Case ROCKWELL_WA2_SYNTH
                         sStr = "ROCKWELL_WA2_SYNTH"
                    Case ROCKWELL_WA2_MIXER
                         sStr = "ROCKWELL_WA2_MIXER"
                    Case ROCKWELL_WA2_MPU401_IN
                         sStr = "ROCKWELL_WA2_MPU401_IN"
                    Case ROCKWELL_WA2_MPU401_OUT
                         sStr = "ROCKWELL_WA2_MPU401_OUT"
               End Select
          Case MM_XYZ ' Co. XYZ for testing
               sStr = "No Associated Product ID"
          Case MM_OPCODE ' Opcode Systems
               sStr = "No Associated Product ID"
          Case MM_VOXWARE ' Voxware Inc
               sStr = "No Associated Product ID"
          Case MM_NORTHERN_TELECOM ' Northern Telecom Limited
               Select Case lID
                    Case MM_NORTEL_MPXAC_WAVEIN
                         sStr = "MM_NORTEL_MPXAC_WAVEIN"
                    Case MM_NORTEL_MPXAC_WAVEOUT
                         sStr = "MM_NORTEL_MPXAC_WAVEOUT"
               End Select
          Case MM_APICOM ' APICOM
               sStr = "No Associated Product ID"
          Case MM_GRANDE ' Grande Software
               sStr = "No Associated Product ID"
          Case MM_ADDX ' ADDX
               Select Case lID
                    Case MM_ADDX_PCTV_DIGITALMIX
                         sStr = "MM_ADDX_PCTV_DIGITALMIX"
                    Case MM_ADDX_PCTV_WAVEIN
                         sStr = "MM_ADDX_PCTV_WAVEIN"
                    Case MM_ADDX_PCTV_WAVEOUT
                         sStr = "MM_ADDX_PCTV_WAVEOUT"
                    Case MM_ADDX_PCTV_MIXER
                         sStr = "MM_ADDX_PCTV_MIXER"
                    Case MM_ADDX_PCTV_AUX_CD
                         sStr = "MM_ADDX_PCTV_AUX_CD"
                    Case MM_ADDX_PCTV_AUX_LINE
                         sStr = "MM_ADDX_PCTV_AUX_LINE"
               End Select
          Case MM_WILDCAT ' Wildcat Canyon Software
               sStr = "MM_WILDCAT_AUTOSCOREMIDIIN"
          Case MM_RHETOREX ' Rhetorex Inc
               Select Case lID
                    Case MM_RHETOREX_WAVEIN
                         sStr = "MM_RHETOREX_WAVEIN"
                    Case MM_RHETOREX_WAVEOUT
                         sStr = "MM_RHETOREX_WAVEOUT"
               End Select
          Case MM_BROOKTREE ' Brooktree Corporation
               Select Case lID
                    Case MM_BTV_WAVEIN
                         sStr = "MM_BTV_WAVEIN"
                    Case MM_BTV_WAVEOUT
                         sStr = "MM_BTV_WAVEOUT"
                    Case MM_BTV_MIDIIN
                         sStr = "MM_BTV_MIDIIN"
                    Case MM_BTV_MIDIOUT
                         sStr = "MM_BTV_MIDIOUT"
                    Case MM_BTV_MIDISYNTH
                         sStr = "MM_BTV_MIDISYNTH"
                    Case MM_BTV_AUX_LINE
                         sStr = "MM_BTV_AUX_LINE"
                    Case MM_BTV_AUX_MIC
                         sStr = "MM_BTV_AUX_MIC"
                    Case MM_BTV_AUX_CD
                         sStr = "MM_BTV_AUX_CD"
                    Case MM_BTV_DIGITALIN
                         sStr = "MM_BTV_DIGITALIN"
                    Case MM_BTV_DIGITALOUT
                         sStr = "MM_BTV_DIGITALOUT"
                    Case MM_BTV_MIDIWAVESTREAM
                         sStr = "MM_BTV_MIDIWAVESTREAM"
                    Case MM_BTV_MIXER
                         sStr = "MM_BTV_MIXER"
               End Select
          Case MM_ENSONIQ ' ENSONIQ Corporation
               Select Case lID
                    Case MM_ENSONIQ_SOUNDSCAPE
                         sStr = "MM_ENSONIQ_SOUNDSCAPE"
                    Case MM_SOUNDSCAPE_WAVEOUT
                         sStr = "MM_SOUNDSCAPE_WAVEOUT"
                    Case MM_SOUNDSCAPE_WAVEOUT_AUX
                         sStr = "MM_SOUNDSCAPE_WAVEOUT_AUX"
                    Case MM_SOUNDSCAPE_WAVEIN
                         sStr = "MM_SOUNDSCAPE_WAVEIN"
                    Case MM_SOUNDSCAPE_MIDIOUT
                         sStr = "MM_SOUNDSCAPE_MIDIOUT"
                    Case MM_SOUNDSCAPE_MIDIIN
                         sStr = "MM_SOUNDSCAPE_MIDIIN"
                    Case MM_SOUNDSCAPE_SYNTH
                         sStr = "MM_SOUNDSCAPE_SYNTH"
                    Case MM_SOUNDSCAPE_MIXER
                         sStr = "MM_SOUNDSCAPE_MIXER"
                    Case MM_SOUNDSCAPE_AUX
                         sStr = "MM_SOUNDSCAPE_AUX"
               End Select
          Case MM_FAST ' FAST Multimedia AG
               sStr = "No Associated Product ID"
          Case MM_NVIDIA ' NVidia Corporation
               Select Case lID
                    Case MM_NVIDIA_WAVEOUT
                        sStr = "MM_NVIDIA_WAVEOUT"
                    Case MM_NVIDIA_WAVEIN
                         sStr = "MM_NVIDIA_WAVEIN"
                    Case MM_NVIDIA_MIDIOUT
                         sStr = "MM_NVIDIA_MIDIOUT"
                    Case MM_NVIDIA_MIDIIN
                         sStr = "MM_NVIDIA_MIDIIN"
                    Case MM_NVIDIA_GAMEPORT
                         sStr = "MM_NVIDIA_GAMEPORT"
                    Case MM_NVIDIA_MIXER
                         sStr = "MM_NVIDIA_MIXER"
                    Case MM_NVIDIA_AUX
                         sStr = "MM_NVIDIA_AUX"
               End Select
          Case MM_OKSORI ' OKSORI Co., Ltd.
               Select Case lID
                    Case MM_OKSORI_OSR8_WAVEOUT
                         sStr = "MM_OKSORI_OSR8_WAVEOUT"
                    Case MM_OKSORI_OSR8_WAVEIN
                         sStr = "MM_OKSORI_OSR8_WAVEIN"
                    Case MM_OKSORI_OSR16_WAVEOUT
                         sStr = "MM_OKSORI_OSR16_WAVEOUT"
                    Case MM_OKSORI_OSR16_WAVEIN
                         sStr = "MM_OKSORI_OSR16_WAVEIN"
                    Case MM_OKSORI_FM_OPL4
                         sStr = "MM_OKSORI_FM_OPL4"
                    Case MM_OKSORI_MIX_MASTER
                         sStr = "MM_OKSORI_MIX_MASTER"
                    Case MM_OKSORI_MIX_WAVE
                         sStr = "MM_OKSORI_MIX_WAVE"
                    Case MM_OKSORI_MIX_FM
                         sStr = "MM_OKSORI_MIX_FM"
                    Case MM_OKSORI_MIX_LINE
                         sStr = "MM_OKSORI_MIX_LINE"
                    Case MM_OKSORI_MIX_CD
                         sStr = "MM_OKSORI_MIX_CD"
                    Case MM_OKSORI_MIX_MIC
                         sStr = "MM_OKSORI_MIX_MIC"
                    Case MM_OKSORI_MIX_ECHO
                         sStr = "MM_OKSORI_MIX_ECHO"
                    Case MM_OKSORI_MIX_AUX1
                         sStr = "MM_OKSORI_MIX_AUX1"
                    Case MM_OKSORI_MIX_LINE1
                         sStr = "MM_OKSORI_MIX_LINE1"
                    Case MM_OKSORI_EXT_MIC1
                         sStr = "MM_OKSORI_EXT_MIC1"
                    Case MM_OKSORI_EXT_MIC2
                         sStr = "MM_OKSORI_EXT_MIC2"
                    Case MM_OKSORI_MIDIOUT
                         sStr = "MM_OKSORI_MIDIOUT"
                    Case MM_OKSORI_MIDIIN
                         sStr = "MM_OKSORI_MIDIIN"
                    Case MM_OKSORI_MPEG_CDVISION
                         sStr = "MM_OKSORI_MPEG_CDVISION"
               End Select
          Case MM_DIACOUSTICS ' DiAcoustics, Inc.
               sStr = "MM_DIACOUSTICS_DRUM_ACTION"
          Case MM_GULBRANSEN ' Gulbransen, Inc.
               sStr = "No Associated Product ID"
          Case MM_KAY_ELEMETRICS ' Kay Elemetrics, Inc.
               Select Case lID
                    Case MM_KAY_ELEMETRICS_CSL
                         sStr = "MM_KAY_ELEMETRICS_CSL"
                    Case MM_KAY_ELEMETRICS_CSL_DAT
                         sStr = "MM_KAY_ELEMETRICS_CSL_DAT"
                    Case MM_KAY_ELEMETRICS_CSL_4CHANNEL
                         sStr = "MM_KAY_ELEMETRICS_CSL_4CHANNEL"
               End Select
          Case MM_CRYSTAL ' Crystal Semiconductor Corporation
               Select Case lID
                    Case MM_CRYSTAL_CS4232_WAVEIN
                         sStr = "MM_CRYSTAL_CS4232_WAVEIN"
                    Case MM_CRYSTAL_CS4232_WAVEOUT
                         sStr = "MM_CRYSTAL_CS4232_WAVEOUT"
                    Case MM_CRYSTAL_CS4232_WAVEMIXER
                         sStr = "MM_CRYSTAL_CS4232_WAVEMIXER"
                    Case MM_CRYSTAL_CS4232_WAVEAUX_AUX1
                         sStr = "MM_CRYSTAL_CS4232_WAVEAUX_AUX1"
                    Case MM_CRYSTAL_CS4232_WAVEAUX_AUX2
                         sStr = "MM_CRYSTAL_CS4232_WAVEAUX_AUX2"
                    Case MM_CRYSTAL_CS4232_WAVEAUX_LINE
                         sStr = "MM_CRYSTAL_CS4232_WAVEAUX_LINE"
                    Case MM_CRYSTAL_CS4232_WAVEAUX_MONO
                         sStr = "MM_CRYSTAL_CS4232_WAVEAUX_MONO"
                    Case MM_CRYSTAL_CS4232_WAVEAUX_MASTER
                         sStr = "MM_CRYSTAL_CS4232_WAVEAUX_MASTER"
                    Case MM_CRYSTAL_CS4232_MIDIIN
                         sStr = "MM_CRYSTAL_CS4232_MIDIIN"
                    Case MM_CRYSTAL_CS4232_MIDIOUT
                         sStr = "MM_CRYSTAL_CS4232_MIDIOUT"
                    Case MM_CRYSTAL_CS4232_INPUTGAIN_AUX1
                         sStr = "MM_CRYSTAL_CS4232_INPUTGAIN_AUX1"
                    Case MM_CRYSTAL_CS4232_INPUTGAIN_AUX1
                         sStr = "MM_CRYSTAL_CS4232_INPUTGAIN_AUX1"
               End Select
          Case MM_SPLASH_STUDIOS ' Splash Studios
               sStr = "No Associated Product ID"
          Case MM_QUARTERDECK ' Quarterdeck Corporation
               Select Case lID
                    Case MM_QUARTERDECK_LHWAVEIN
                         sStr = "MM_QUARTERDECK_LHWAVEIN"
                    Case MM_QUARTERDECK_LHWAVEOUT
                         sStr = "MM_QUARTERDECK_LHWAVEOUT"
               End Select
          Case MM_TDK ' TDK Corporation
               Select Case lID
                    Case MM_TDK_MW_MIDI_SYNTH
                         sStr = "MM_TDK_MW_MIDI_SYNTH"
                    Case MM_TDK_MW_MIDI_IN
                         sStr = "MM_TDK_MW_MIDI_IN"
                    Case MM_TDK_MW_MIDI_OUT
                         sStr = "MM_TDK_MW_MIDI_OUT"
                    Case MM_TDK_MW_WAVE_IN
                         sStr = "MM_TDK_MW_WAVE_IN"
                    Case MM_TDK_MW_WAVE_OUT
                         sStr = "MM_TDK_MW_WAVE_OUT"
                    Case MM_TDK_MW_AUX
                         sStr = "MM_TDK_MW_AUX"
                    Case MM_TDK_MW_MIXER
                         sStr = "MM_TDK_MW_MIXER"
                    Case MM_TDK_MW_AUX_MASTER
                         sStr = "MM_TDK_MW_AUX_MASTER"
                    Case MM_TDK_MW_AUX_BASS
                         sStr = "MM_TDK_MW_AUX_BASS"
                    Case MM_TDK_MW_AUX_TREBLE
                         sStr = "MM_TDK_MW_AUX_TREBLE"
                    Case MM_TDK_MW_AUX_MIDI_VOL
                         sStr = "MM_TDK_MW_AUX_MIDI_VOL"
                    Case MM_TDK_MW_AUX_WAVE_VOL
                         sStr = "MM_TDK_MW_AUX_WAVE_VOL"
                    Case MM_TDK_MW_AUX_WAVE_RVB
                         sStr = "MM_TDK_MW_AUX_WAVE_RVB"
                    Case MM_TDK_MW_AUX_WAVE_CHR
                         sStr = "MM_TDK_MW_AUX_WAVE_CHR"
                    Case MM_TDK_MW_AUX_VOL
                         sStr = "MM_TDK_MW_AUX_VOL"
                    Case MM_TDK_MW_AUX_RVB
                         sStr = "MM_TDK_MW_AUX_RVB"
                    Case MM_TDK_MW_AUX_CHR
                         sStr = "MM_TDK_MW_AUX_CHR"
               End Select
          Case MM_DIGITAL_AUDIO_LABS ' Digital Audio Labs, Inc.
               Select Case lID
                    Case MM_DIGITAL_AUDIO_LABS_V8
                         sStr = "MM_DIGITAL_AUDIO_LABS_V8"
                    Case MM_DIGITAL_AUDIO_LABS_CPRO
                         sStr = "MM_DIGITAL_AUDIO_LABS_CPRO"
               End Select
          Case MM_SEERSYS ' Seer Systems, Inc.
               Select Case lID
                    Case MM_SEERSYS_SEERSYNTH
                         sStr = "MM_SEERSYS_SEERSYNTH"
                    Case MM_SEERSYS_SEERWAVE
                         sStr = "MM_SEERSYS_SEERWAVE"
                    Case MM_SEERSYS_SEERMIX
                         sStr = "MM_SEERSYS_SEERMIX"
               End Select
          Case MM_PICTURETEL ' PictureTel Corporation
               sStr = "No Associated Product ID"
          Case MM_ATT_MICROELECTRONICS ' AT&T Microelectronics
               sStr = "No Associated Product ID"
          Case MM_OSPREY ' Osprey Technologies, Inc.
               Select Case lID
                    Case MM_OSPREY_1000WAVEIN
                         sStr = "MM_OSPREY_1000WAVEIN"
                    Case MM_OSPREY_1000WAVEOUT
                         sStr = "MM_OSPREY_1000WAVEOUT"
               End Select
          Case MM_MEDIATRIX ' Mediatrix Peripherals
               sStr = "No Associated Product ID"
          Case MM_SOUNDESIGNS ' SounDesignS M.C.S. Ltd.
               Select Case lID
                    Case MM_SOUNDESIGNS_WAVEIN
                         sStr = "MM_SOUNDESIGNS_WAVEIN"
                    Case MM_SOUNDESIGNS_WAVEOUT
                         sStr = "MM_SOUNDESIGNS_WAVEOUT"
               End Select
          Case MM_ALDIGITAL ' A.L. Digital Ltd.
               sStr = "No Associated Product ID"
          Case MM_SPECTRUM_SIGNAL_PROCESSING ' Spectrum Signal Processing, Inc.
               Select Case lID
                    Case MM_SSP_SNDFESWAVEIN
                         sStr = "MM_SSP_SNDFESWAVEIN"
                    Case MM_SSP_SNDFESWAVEOUT
                         sStr = "MM_SSP_SNDFESWAVEOUT"
                    Case MM_SSP_SNDFESMIDIIN
                         sStr = "MM_SSP_SNDFESMIDIIN"
                    Case MM_SSP_SNDFESMIDIOUT
                         sStr = "MM_SSP_SNDFESMIDIOUT"
                    Case MM_SSP_SNDFESSYNTH
                         sStr = "MM_SSP_SNDFESSYNTH"
                    Case MM_SSP_SNDFESMIX
                         sStr = "MM_SSP_SNDFESMIX"
                    Case MM_SSP_SNDFESAUX
                         sStr = "MM_SSP_SNDFESAUX"
               End Select
          Case MM_ECS ' Electronic Courseware Systems, Inc.
               Select Case lID
                    Case MM_ECS_AADF_MIDI_IN
                         sStr = "MM_ECS_AADF_MIDI_IN"
                    Case MM_ECS_AADF_MIDI_OUT
                         sStr = "MM_ECS_AADF_MIDI_OUT"
                    Case MM_ECS_AADF_WAVE2MIDI_IN
                         sStr = "MM_ECS_AADF_WAVE2MIDI_IN"
               End Select
          Case MM_AMD ' AMD
               Select Case lID
                    Case MM_AMD_INTERWAVE_WAVEIN
                         sStr = "MM_AMD_INTERWAVE_WAVEIN"
                    Case MM_AMD_INTERWAVE_WAVEOUT
                         sStr = "MM_AMD_INTERWAVE_WAVEOUT"
                    Case MM_AMD_INTERWAVE_SYNTH
                         sStr = "MM_AMD_INTERWAVE_SYNTH"
                    Case MM_AMD_INTERWAVE_MIXER1
                         sStr = "MM_AMD_INTERWAVE_MIXER1"
                    Case MM_AMD_INTERWAVE_MIXER2
                         sStr = "MM_AMD_INTERWAVE_MIXER2"
                    Case MM_AMD_INTERWAVE_JOYSTICK
                         sStr = "MM_AMD_INTERWAVE_JOYSTICK"
                    Case MM_AMD_INTERWAVE_EX_CD
                         sStr = "MM_AMD_INTERWAVE_EX_CD"
                    Case MM_AMD_INTERWAVE_MIDIIN
                         sStr = "MM_AMD_INTERWAVE_MIDIIN"
                    Case MM_AMD_INTERWAVE_MIDIOUT
                         sStr = "MM_AMD_INTERWAVE_MIDIOUT"
                    Case MM_AMD_INTERWAVE_AUX1
                         sStr = "MM_AMD_INTERWAVE_AUX1"
                    Case MM_AMD_INTERWAVE_AUX2
                         sStr = "MM_AMD_INTERWAVE_AUX2"
                    Case MM_AMD_INTERWAVE_AUX_MIC
                         sStr = "MM_AMD_INTERWAVE_AUX_MIC"
                    Case MM_AMD_INTERWAVE_AUX_CD
                         sStr = "MM_AMD_INTERWAVE_AUX_CD"
                    Case MM_AMD_INTERWAVE_MONO_IN
                         sStr = "MM_AMD_INTERWAVE_MONO_IN"
                    Case MM_AMD_INTERWAVE_MONO_OUT
                         sStr = "MM_AMD_INTERWAVE_MONO_OUT"
                    Case MM_AMD_INTERWAVE_EX_TELEPHONY
                         sStr = "MM_AMD_INTERWAVE_EX_TELEPHONY"
                    Case MM_AMD_INTERWAVE_WAVEOUT_BASE
                         sStr = "MM_AMD_INTERWAVE_WAVEOUT_BASE"
                    Case MM_AMD_INTERWAVE_WAVEOUT_TREBLE
                         sStr = "MM_AMD_INTERWAVE_WAVEOUT_TREBLE"
                    Case MM_AMD_INTERWAVE_STEREO_ENHANCED
                         sStr = "MM_AMD_INTERWAVE_STEREO_ENHANCED"
               End Select
          Case MM_COREDYNAMICS ' Core Dynamics
               Select Case lID
                    Case MM_COREDYNAMICS_DYNAMIXHR
                         sStr = "MM_COREDYNAMICS_DYNAMIXHR"
                    Case MM_COREDYNAMICS_DYNASONIX_SYNTH
                         sStr = "MM_COREDYNAMICS_DYNASONIX_SYNTH"
                    Case MM_COREDYNAMICS_DYNASONIX_MIDI_IN
                         sStr = "MM_COREDYNAMICS_DYNASONIX_MIDI_IN"
                    Case MM_COREDYNAMICS_DYNASONIX_MIDI_OUT
                         sStr = "MM_COREDYNAMICS_DYNASONIX_MIDI_OUT"
                    Case MM_COREDYNAMICS_DYNASONIX_WAVE_IN
                         sStr = "MM_COREDYNAMICS_DYNASONIX_WAVE_IN"
                    Case MM_COREDYNAMICS_DYNASONIX_WAVE_OUT
                         sStr = "MM_COREDYNAMICS_DYNASONIX_WAVE_OUT"
                    Case MM_COREDYNAMICS_DYNASONIX_AUDIO_IN
                         sStr = "MM_COREDYNAMICS_DYNASONIX_AUDIO_IN"
                    Case MM_COREDYNAMICS_DYNASONIX_AUDIO_OUT
                         sStr = "MM_COREDYNAMICS_DYNASONIX_AUDIO_OUT"
                    Case MM_COREDYNAMICS_DYNAGRAFX_VGA
                         sStr = "MM_COREDYNAMICS_DYNAGRAFX_VGA"
                    Case MM_COREDYNAMICS_DYNAGRAFX_WAVE_IN
                         sStr = "MM_COREDYNAMICS_DYNAGRAFX_WAVE_IN"
                    Case MM_COREDYNAMICS_DYNAGRAFX_WAVE_OUT
                         sStr = "MM_COREDYNAMICS_DYNAGRAFX_WAVE_OUT"
               End Select
          Case MM_CANAM ' CANAM Computers
               Select Case lID
                    Case MM_CANAM_CBXWAVEOUT
                         sStr = "MM_CANAM_CBXWAVEOUT"
                    Case MM_CANAM_CBXWAVEIN
                         sStr = "MM_CANAM_CBXWAVEIN"
               End Select
          Case MM_SOFTSOUND ' Softsound, Ltd.
               sStr = "MM_SOFTSOUND_CODEC"
          Case MM_NORRIS ' Norris Communications, Inc.
               sStr = "MM_NORRIS_VOICELINK"
          Case MM_DDD ' Danka Data Devices
               Select Case lID
                    Case MM_DDD_MIDILINK_MIDIIN
                         sStr = "MM_DDD_MIDILINK_MIDIIN"
                    Case MM_DDD_MIDILINK_MIDIOUT
                         sStr = "MM_DDD_MIDILINK_MIDIOUT"
               End Select
          Case MM_EUPHONICS ' EuPhonics
               sStr = "No Associated Product ID"
          Case MM_PRECEPT ' Precept Software, Inc.
               sStr = "No Associated Product ID"
          Case MM_CRYSTAL_NET ' Crystal Net Corporation
               sStr = "No Associated Product ID"
          Case MM_CHROMATIC ' Chromatic Research, Inc
               sStr = "No Associated Product ID"
          Case MM_VOICEINFO ' Voice Information Systems, Inc
               sStr = "No Associated Product ID"
          Case MM_VIENNASYS ' Vienna Systems
               sStr = "MM_VIENNASYS_TSP_WAVE_DRIVER"
          Case MM_CONNECTIX ' Connectix Corporation
               sStr = "No Associated Product ID"
          Case MM_GADGETLABS ' Gadget Labs LLC
               Select Case lID
                    Case MM_GADGETLABS_WAVE44_WAVEIN
                         sStr = "MM_GADGETLABS_WAVE44_WAVEIN"
                    Case MM_GADGETLABS_WAVE44_WAVEOUT
                         sStr = "MM_GADGETLABS_WAVE44_WAVEOUT"
                    Case MM_GADGETLABS_WAVE42_WAVEIN
                         sStr = "MM_GADGETLABS_WAVE42_WAVEIN"
                    Case MM_GADGETLABS_WAVE42_WAVEOUT
                         sStr = "MM_GADGETLABS_WAVE42_WAVEOUT"
                    Case MM_GADGETLABS_WAVE4_MIDIIN
                         sStr = "MM_GADGETLABS_WAVE4_MIDIIN"
                    Case MM_GADGETLABS_WAVE4_MIDIOUT
                         sStr = "MM_GADGETLABS_WAVE4_MIDIOUT"
               End Select
          Case MM_FRONTIER ' Frontier Design Group LLC
               Select Case lID
                    Case MM_FRONTIER_WAVECENTER_MIDIIN
                         sStr = "MM_FRONTIER_WAVECENTER_MIDIIN"
                    Case MM_FRONTIER_WAVECENTER_MIDIOUT
                         sStr = "MM_FRONTIER_WAVECENTER_MIDIOUT"
                    Case MM_FRONTIER_WAVECENTER_WAVEIN
                         sStr = "MM_FRONTIER_WAVECENTER_WAVEIN"
                    Case MM_FRONTIER_WAVECENTER_WAVEOUT
                         sStr = "MM_FRONTIER_WAVECENTER_WAVEOUT"
               End Select
          Case MM_VIONA ' Viona Development GmbH
               Select Case lID
                    Case MM_VIONA_QVINPCI_MIXER
                         sStr = "MM_VIONA_QVINPCI_MIXER"
                    Case MM_VIONA_QVINPCI_WAVEIN
                         sStr = "MM_VIONA_QVINPCI_WAVEIN"
                    Case MM_VIONA_QVINPCI_WAVEOUT
                         sStr = "MM_VIONA_QVINPCI_WAVEOUT"
                    Case MM_VIONA_BUSTER_MIXER
                         sStr = "MM_VIONA_BUSTER_MIXER"
                    Case MM_VIONA_CINEMASTER_MIXER
                         sStr = "MM_VIONA_CINEMASTER_MIXER"
                    Case MM_VIONA_CONCERTO_MIXER
                         sStr = "MM_VIONA_CONCERTO_MIXER"
               End Select
          Case MM_CASIO ' Casio Computer Co., LTD
               Select Case lID
                    Case MM_CASIO_WP150_MIDIOUT
                         sStr = "MM_CASIO_WP150_MIDIOUT"
                    Case MM_CASIO_WP150_MIDIIN
                         sStr = "MM_CASIO_WP150_MIDIIN"
               End Select
          Case MM_DIAMONDMM ' Diamond Multimedia
               Select Case lID
                    Case MM_DIMD_PLATFORM
                         sStr = "MM_DIMD_PLATFORM"
                    Case MM_DIMD_DIRSOUND
                         sStr = "MM_DIMD_DIRSOUND"
                    Case MM_DIMD_VIRTMPU
                         sStr = "MM_DIMD_VIRTMPU"
                    Case MM_DIMD_VIRTSB
                         sStr = "MM_DIMD_VIRTSB"
                    Case MM_DIMD_VIRTJOY
                         sStr = "MM_DIMD_VIRTJOY"
                    Case MM_DIMD_WAVEIN
                         sStr = "MM_DIMD_WAVEIN"
                    Case MM_DIMD_WAVEOUT
                         sStr = "MM_DIMD_WAVEOUT"
                    Case MM_DIMD_MIDIIN
                         sStr = "MM_DIMD_MIDIIN"
                    Case MM_DIMD_MIDIOUT
                         sStr = "MM_DIMD_MIDIOUT"
                    Case MM_DIMD_AUX_LINE
                         sStr = "MM_DIMD_AUX_LINE"
                    Case MM_DIMD_MIXER
                         sStr = "MM_DIMD_MIXER"
               End Select
          Case MM_S3 ' S3
               Select Case lID
                    Case MM_S3_WAVEOUT
                         sStr = "MM_S3_WAVEOUT"
                    Case MM_S3_WAVEIN
                         sStr = "MM_S3_WAVEIN"
                    Case MM_S3_MIDIOUT
                         sStr = "MM_S3_MIDIOUT"
                    Case MM_S3_MIDIIN
                         sStr = "MM_S3_MIDIIN"
                    Case MM_S3_FMSYNTH
                         sStr = "MM_S3_FMSYNTH"
                    Case MM_S3_MIXER
                         sStr = "MM_S3_MIXER"
                    Case MM_S3_AUX
                         sStr = "MM_S3_AUX"
               End Select
          Case MM_FRAUNHOFER_IIS ' Fraunhofer
               sStr = "MM_FHGIIS_MPEGLAYER3"
     End Select
     If Len(sStr) = False Then sStr = "No Associated Product ID"
     ' Return function
     GetProductID = sStr
End Function ' GetProductID




