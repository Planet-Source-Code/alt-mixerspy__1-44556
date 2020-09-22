Attribute VB_Name = "mdlManufacturerID"
Option Explicit
'**************************************************************************************************
'  Copyright © 2003, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************
' manufacturer IDs taken from MMREG.H
Public Const MM_MICROSOFT = 1                            '  Microsoft Corporation
Public Const MM_CREATIVE = 2                             '  Creative Labs, Inc
Public Const MM_MEDIAVISION = 3                          '  Media Vision, Inc.
Public Const MM_FUJITSU = 4                              '  Fujitsu Corp.
Public Const MM_ARTISOFT = 20                            '  Artisoft, Inc.
Public Const MM_TURTLE_BEACH = 21                        '  Turtle Beach, Inc.
Public Const MM_IBM = 22                                 '  IBM Corporation
Public Const MM_VOCALTEC = 23                            '  Vocaltec LTD.
Public Const MM_ROLAND = 24                              '  Roland
Public Const MM_DSP_SOLUTIONS = 25                       '  DSP Solutions, Inc.
Public Const MM_NEC = 26                                 '  NEC
Public Const MM_ATI = 27                                 '  ATI
Public Const MM_WANGLABS = 28                            '  Wang Laboratories, Inc
Public Const MM_TANDY = 29                               '  Tandy Corporation
Public Const MM_VOYETRA = 30                             '  Voyetra
Public Const MM_ANTEX = 31                               '  Antex Electronics Corporation
Public Const MM_ICL_PS = 32                              '  ICL Personal Systems
Public Const MM_INTEL = 33                               '  Intel Corporation
Public Const MM_GRAVIS = 34                              '  Advanced Gravis
Public Const MM_VAL = 35                                 '  Video Associates Labs, Inc.
Public Const MM_INTERACTIVE = 36                         '  InterActive Inc
Public Const MM_YAMAHA = 37                              '  Yamaha Corporation of America
Public Const MM_EVEREX = 38                              '  Everex Systems, Inc
Public Const MM_ECHO = 39                                '  Echo Speech Corporation
Public Const MM_SIERRA = 40                              '  Sierra Semiconductor Corp
Public Const MM_CAT = 41                                 '  Computer Aided Technologies
Public Const MM_APPS = 42                                '  APPS Software International
Public Const MM_DSP_GROUP = 43                           '  DSP Group, Inc
Public Const MM_MELABS = 44                              '  microEngineering Labs
Public Const MM_COMPUTER_FRIENDS = 45                    '  Computer Friends, Inc.
Public Const MM_ESS = 46                                 '  ESS Technology
Public Const MM_AUDIOFILE = 47                           '  Audio, Inc.
Public Const MM_MOTOROLA = 48                            '  Motorola, Inc.
Public Const MM_CANOPUS = 49                             '  Canopus, co., Ltd.
Public Const MM_EPSON = 50                               '  Seiko Epson Corporation
Public Const MM_TRUEVISION = 51                          '  Truevision
Public Const MM_AZTECH = 52                              '  Aztech Labs, Inc.
Public Const MM_VIDEOLOGIC = 53                          '  Videologic
Public Const MM_SCALACS = 54                             '  SCALACS
Public Const MM_KORG = 55                                '  Korg Inc.
Public Const MM_APT = 56                                 '  Audio Processing Technology
Public Const MM_ICS = 57                                 '  Integrated Circuit Systems, Inc.
Public Const MM_ITERATEDSYS = 58                         '  Iterated Systems, Inc.
Public Const MM_METHEUS = 59                             '  Metheus
Public Const MM_LOGITECH = 60                            '  Logitech, Inc.
Public Const MM_WINNOV = 61                              '  Winnov, Inc.
Public Const MM_NCR = 62                                 '  NCR Corporation
Public Const MM_EXAN = 63                                '  EXAN
Public Const MM_AST = 64                                 '  AST Research Inc.
Public Const MM_WILLOWPOND = 65                          '  Willow Pond Corporation
Public Const MM_SONICFOUNDRY = 66                        '  Sonic Foundry
Public Const MM_VITEC = 67                               '  Vitec Multimedia
Public Const MM_MOSCOM = 68                              '  MOSCOM Corporation
Public Const MM_SILICONSOFT = 69                         '  Silicon Soft, Inc.
Public Const MM_SUPERMAC = 73                            '  Supermac
Public Const MM_AUDIOPT = 74                             '  Audio Processing Technology
Public Const MM_SPEECHCOMP = 76                          '  Speech Compression
Public Const MM_AHEAD = 77                               '  Ahead, Inc.
Public Const MM_DOLBY = 78                               '  Dolby Laboratories
Public Const MM_OKI = 79                                 '  OKI
Public Const MM_AURAVISION = 80                          '  AuraVision Corporation
Public Const MM_OLIVETTI = 81                            '  Ing C. Olivetti & C., S.p.A.
Public Const MM_IOMAGIC = 82                             '  I/O Magic Corporation
Public Const MM_MATSUSHITA = 83                          '  Matsushita Electric Industrial Co., LTD.
Public Const MM_CONTROLRES = 84                          '  Control Resources Limited
Public Const MM_XEBEC = 85                               '  Xebec Multimedia Solutions Limited
Public Const MM_NEWMEDIA = 86                            '  New Media Corporation
Public Const MM_NMS = 87                                 '  Natural MicroSystems
Public Const MM_LYRRUS = 88                              '  Lyrrus Inc.
Public Const MM_COMPUSIC = 89                            '  Compusic
Public Const MM_OPTI = 90                                '  OPTi Computers Inc.
Public Const MM_ADLACC = 91                              '  Adlib Accessories Inc.
Public Const MM_COMPAQ = 92                              '  Compaq Computer Corp.
Public Const MM_DIALOGIC = 93                            '  Dialogic Corporation
Public Const MM_INSOFT = 94                              '  InSoft, Inc.
Public Const MM_MPTUS = 95                               '  M.P. Technologies, Inc.
Public Const MM_WEITEK = 96                              '  Weitek
Public Const MM_LERNOUT_AND_HAUSPIE = 97                 '  Lernout & Hauspie
Public Const MM_QCIAR = 98                               '  Quanta Computer Inc.
Public Const MM_APPLE = 99                               '  Apple Computer, Inc.
Public Const MM_DIGITAL = 100                            '  Digital Equipment Corporation
Public Const MM_MOTU = 101                               '  Mark of the Unicorn
Public Const MM_WORKBIT = 102                            '  Workbit Corporation
Public Const MM_OSITECH = 103                            '  Ositech Communications Inc.
Public Const MM_MIRO = 104                               '  miro Computer Products AG
Public Const MM_CIRRUSLOGIC = 105                        '  Cirrus Logic
Public Const MM_ISOLUTION = 106                          '  ISOLUTION  B.V.
Public Const MM_HORIZONS = 107                           '  Horizons Technology, Inc
Public Const MM_CONCEPTS = 108                           '  Computer Concepts Ltd
Public Const MM_VTG = 109                                '  Voice Technologies Group, Inc.
Public Const MM_RADIUS = 110                             '  Radius
Public Const MM_ROCKWELL = 111                           '  Rockwell International
Public Const MM_XYZ = 112                                '  Co. XYZ for testing
Public Const MM_OPCODE = 113                             '  Opcode Systems
Public Const MM_VOXWARE = 114                            '  Voxware Inc
Public Const MM_NORTHERN_TELECOM = 115                   '  Northern Telecom Limited
Public Const MM_APICOM = 116                             '  APICOM
Public Const MM_GRANDE = 117                             '  Grande Software
Public Const MM_ADDX = 118                               '  ADDX
Public Const MM_WILDCAT = 119                            '  Wildcat Canyon Software
Public Const MM_RHETOREX = 120                           '  Rhetorex Inc
Public Const MM_BROOKTREE = 121                          '  Brooktree Corporation
Public Const MM_ENSONIQ = 125                            '  ENSONIQ Corporation
Public Const MM_FAST = 126                               '  ///FAST Multimedia AG
Public Const MM_NVIDIA = 127                             '  NVidia Corporation
Public Const MM_OKSORI = 128                             '  OKSORI Co., Ltd.
Public Const MM_DIACOUSTICS = 129                        '  DiAcoustics, Inc.
Public Const MM_GULBRANSEN = 130                         '  Gulbransen, Inc.
Public Const MM_KAY_ELEMETRICS = 131                     '  Kay Elemetrics, Inc.
Public Const MM_CRYSTAL = 132                            '  Crystal Semiconductor Corporation
Public Const MM_SPLASH_STUDIOS = 133                     '  Splash Studios
Public Const MM_QUARTERDECK = 134                        '  Quarterdeck Corporation
Public Const MM_TDK = 135                                '  TDK Corporation
Public Const MM_DIGITAL_AUDIO_LABS = 136                 '  Digital Audio Labs, Inc.
Public Const MM_SEERSYS = 137                            '  Seer Systems, Inc.
Public Const MM_PICTURETEL = 138                         '  PictureTel Corporation
Public Const MM_ATT_MICROELECTRONICS = 139               '  AT&T Microelectronics
Public Const MM_OSPREY = 140                             '  Osprey Technologies, Inc.
Public Const MM_MEDIATRIX = 141                          '  Mediatrix Peripherals
Public Const MM_SOUNDESIGNS = 142                        '  SounDesignS M.C.S. Ltd.
Public Const MM_ALDIGITAL = 143                          '  A.L. Digital Ltd.
Public Const MM_SPECTRUM_SIGNAL_PROCESSING = 144         '  Spectrum Signal Processing, Inc.
Public Const MM_ECS = 145                                '  Electronic Courseware Systems, Inc.
Public Const MM_AMD = 146                                '  AMD
Public Const MM_COREDYNAMICS = 147                       '  Core Dynamics
Public Const MM_CANAM = 148                              '  CANAM Computers
Public Const MM_SOFTSOUND = 149                          '  Softsound, Ltd.
Public Const MM_NORRIS = 150                             '  Norris Communications, Inc.
Public Const MM_DDD = 151                                '  Danka Data Devices
Public Const MM_EUPHONICS = 152                          '  EuPhonics
Public Const MM_PRECEPT = 153                            '  Precept Software, Inc.
Public Const MM_CRYSTAL_NET = 154                        '  Crystal Net Corporation
Public Const MM_CHROMATIC = 155                          '  Chromatic Research, Inc
Public Const MM_VOICEINFO = 156                          '  Voice Information Systems, Inc
Public Const MM_VIENNASYS = 157                          '  Vienna Systems
Public Const MM_CONNECTIX = 158                          '  Connectix Corporation
Public Const MM_GADGETLABS = 159                         '  Gadget Labs LLC
Public Const MM_FRONTIER = 160                           '  Frontier Design Group LLC
Public Const MM_VIONA = 161                              '  Viona Development GmbH
Public Const MM_CASIO = 162                              '  Casio Computer Co., LTD
Public Const MM_DIAMONDMM = 163                          '  Diamond Multimedia
Public Const MM_S3 = 164                                 '  S3
Public Const MM_FRAUNHOFER_IIS = 172                     '  Fraunhofer

Public Function GetManufacturer(ByVal lID As Long) As String
     Dim sStr As String
     Select Case lID
          Case MM_MICROSOFT
               sStr = "MM_MICROSOFT"
          Case MM_CREATIVE
               sStr = "MM_CREATIVE"
          Case MM_MEDIAVISION
               sStr = "MM_MEDIAVISION"
          Case MM_FUJITSU
               sStr = "MM_FUJITSU"
          Case MM_ARTISOFT
               sStr = "MM_ARTISOFT"
          Case MM_TURTLE_BEACH
               sStr = "MM_TURTLE_BEACH"
          Case MM_IBM
               sStr = "MM_IBM"
          Case MM_VOCALTEC
               sStr = "MM_VOCALTEC"
          Case MM_ROLAND
               sStr = "MM_ROLAND"
          Case MM_DSP_SOLUTIONS
               sStr = "MM_DSP_SOLUTIONS"
          Case MM_NEC
               sStr = "MM_NEC"
          Case MM_ATI
               sStr = "MM_ATI"
          Case MM_WANGLABS
               sStr = "MM_WANGLABS"
          Case MM_TANDY
               sStr = "MM_TANDY"
          Case MM_VOYETRA
               sStr = "MM_VOYETRA"
          Case MM_ANTEX
               sStr = "MM_ANTEX"
          Case MM_ICL_PS
               sStr = "MM_ICL_PS"
          Case MM_INTEL
               sStr = "MM_INTEL"
          Case MM_GRAVIS
               sStr = "MM_GRAVIS"
          Case MM_VAL
               sStr = "MM_VAL"
          Case MM_INTERACTIVE
               sStr = "MM_INTERACTIVE"
          Case MM_YAMAHA
               sStr = "MM_YAMAHA"
          Case MM_EVEREX
               sStr = "MM_EVEREX"
          Case MM_ECHO
               sStr = "MM_ECHO"
          Case MM_SIERRA
               sStr = "MM_SIERRA"
          Case MM_CAT
               sStr = "MM_CAT"
          Case MM_APPS
               sStr = "MM_APPS"
          Case MM_DSP_GROUP
               sStr = "MM_DSP_GROUP"
          Case MM_COMPUTER_FRIENDS
               sStr = "MM_COMPUTER_FRIENDS"
          Case MM_ESS
               sStr = "MM_ESS"
          Case MM_AUDIOFILE
               sStr = "MM_AUDIOFILE"
          Case MM_MOTOROLA
               sStr = "MM_MOTOROLA"
          Case MM_CANOPUS
               sStr = "MM_CANOPUS"
          Case MM_EPSON
               sStr = "MM_EPSON"
          Case MM_TRUEVISION
               sStr = "MM_TRUEVISION"
          Case MM_AZTECH
               sStr = "MM_AZTECH"
          Case MM_VIDEOLOGIC
               sStr = "MM_VIDEOLOGIC"
          Case MM_SCALACS
               sStr = "MM_SCALACS"
          Case MM_KORG
               sStr = "MM_KORG"
          Case MM_APT
               sStr = "MM_APT"
          Case MM_ICS
               sStr = "MM_ICS"
          Case MM_ITERATEDSYS
               sStr = "MM_ITERATEDSYS"
          Case MM_METHEUS
               sStr = "MM_METHEUS"
          Case MM_LOGITECH
               sStr = "MM_LOGITECH"
          Case MM_WINNOV
               sStr = "MM_WINNOV"
          Case MM_NCR
               sStr = "MM_NCR"
          Case MM_EXAN
               sStr = "MM_EXAN"
          Case MM_AST
               sStr = "MM_AST"
          Case MM_WILLOWPOND
               sStr = "MM_WILLOWPOND"
          Case MM_SONICFOUNDRY
               sStr = "MM_SONICFOUNDRY"
          Case MM_VITEC
               sStr = "MM_VITEC"
          Case MM_MOSCOM
               sStr = "MM_MOSCOM"
          Case MM_SILICONSOFT
               sStr = "MM_SILICONSOFT"
          Case MM_SUPERMAC
               sStr = "MM_SUPERMAC"
          Case MM_AUDIOPT
               sStr = "MM_AUDIOPT"
          Case MM_SPEECHCOMP
               sStr = "MM_SPEECHCOMP"
          Case MM_AHEAD
               sStr = "MM_AHEAD"
          Case MM_DOLBY
               sStr = "MM_DOLBY"
          Case MM_OKI
               sStr = "MM_OKI"
          Case MM_AURAVISION
               sStr = "MM_AURAVISION"
          Case MM_OLIVETTI
               sStr = "MM_OLIVETTI"
          Case MM_IOMAGIC
               sStr = "MM_IOMAGIC"
          Case MM_MATSUSHITA
               sStr = "MM_MATSUSHITA"
          Case MM_CONTROLRES
               sStr = "MM_CONTROLRES"
          Case MM_XEBEC
               sStr = "MM_XEBEC"
          Case MM_NEWMEDIA
               sStr = "MM_NEWMEDIA"
          Case MM_NMS
               sStr = "MM_NMS"
          Case MM_LYRRUS
               sStr = "MM_LYRRUS"
          Case MM_COMPUSIC
               sStr = "MM_COMPUSIC"
          Case MM_OPTI
               sStr = "MM_OPTI"
          Case MM_ADLACC
               sStr = "MM_ADLACC"
          Case MM_COMPAQ
               sStr = "MM_COMPAQ"
          Case MM_DIALOGIC
               sStr = "MM_DIALOGIC"
          Case MM_INSOFT
               sStr = "MM_INSOFT"
          Case MM_MPTUS
               sStr = "MM_MPTUS"
          Case MM_WEITEK
               sStr = "MM_WEITEK"
          Case MM_LERNOUT_AND_HAUSPIE
               sStr = "MM_LERNOUT_AND_HAUSPIE"
          Case MM_QCIAR
               sStr = "MM_QCIAR"
          Case MM_APPLE
               sStr = "MM_APPLE"
          Case MM_DIGITAL
               sStr = "MM_DIGITAL"
          Case MM_MOTU
               sStr = "MM_MOTU"
          Case MM_WORKBIT
               sStr = "MM_WORKBIT"
          Case MM_OSITECH
               sStr = "MM_OSITECH"
          Case MM_MIRO
               sStr = "MM_MIRO"
          Case MM_CIRRUSLOGIC
               sStr = "MM_CIRRUSLOGIC"
          Case MM_ISOLUTION
               sStr = "MM_ISOLUTION"
          Case MM_HORIZONS
               sStr = "MM_HORIZONS"
          Case MM_CONCEPTS
               sStr = "MM_CONCEPTS"
          Case MM_VTG
               sStr = "MM_VTG"
          Case MM_RADIUS
               sStr = "MM_RADIUS"
          Case MM_ROCKWELL
               sStr = "MM_ROCKWELL"
          Case MM_OPCODE
               sStr = "MM_OPCODE"
          Case MM_VOXWARE
               sStr = "MM_VOXWARE"
          Case MM_NORTHERN_TELECOM
               sStr = "MM_NORTHERN_TELECOM"
          Case MM_APICOM
               sStr = "MM_APICOM"
          Case MM_GRANDE
               sStr = "MM_GRANDE"
          Case MM_ADDX
               sStr = "MM_ADDX"
          Case MM_WILDCAT
               sStr = "MM_WILDCAT"
          Case MM_RHETOREX
               sStr = "MM_RHETOREX"
          Case MM_BROOKTREE
               sStr = "MM_BROOKTREE"
          Case MM_ENSONIQ
               sStr = "MM_ENSONIQ"
          Case MM_FAST
               sStr = "MM_FAST"
          Case MM_NVIDIA
               sStr = "MM_NVIDIA"
          Case MM_OKSORI
               sStr = "MM_OKSORI"
          Case MM_DIACOUSTICS
               sStr = "MM_DIACOUSTICS"
          Case MM_GULBRANSEN
               sStr = "MM_GULBRANSEN"
          Case MM_KAY_ELEMETRICS
               sStr = "MM_KAY_ELEMETRICS"
          Case MM_CRYSTAL
               sStr = "MM_CRYSTAL"
          Case MM_SPLASH_STUDIOS
               sStr = "MM_SPLASH_STUDIOS"
          Case MM_QUARTERDECK
               sStr = "MM_QUARTERDECK"
          Case MM_TDK
               sStr = "MM_TDK"
          Case MM_DIGITAL_AUDIO_LABS
               sStr = "MM_DIGITAL_AUDIO_LABS"
          Case MM_SEERSYS
               sStr = "MM_SEERSYS"
          Case MM_PICTURETEL
               sStr = "MM_PICTURETEL"
          Case MM_ATT_MICROELECTRONICS
               sStr = "MM_ATT_MICROELECTRONICS"
          Case MM_OSPREY
               sStr = "MM_OSPREY"
          Case MM_MEDIATRIX
               sStr = "MM_MEDIATRIX"
          Case MM_SOUNDESIGNS
               sStr = "MM_SOUNDESIGNS"
          Case MM_ALDIGITAL
               sStr = "MM_ALDIGITAL"
          Case MM_SPECTRUM_SIGNAL_PROCESSING
               sStr = "MM_SPECTRUM_SIGNAL_PROCESSING"
          Case MM_ECS
               sStr = "MM_ECS"
          Case MM_AMD
               sStr = "MM_AMD"
          Case MM_COREDYNAMICS
               sStr = "MM_COREDYNAMICS"
          Case MM_CANAM
               sStr = "MM_CANAM"
          Case MM_SOFTSOUND
               sStr = "MM_SOFTSOUND"
          Case MM_NORRIS
               sStr = "MM_NORRIS"
          Case MM_DDD
               sStr = "MM_DDD"
          Case MM_EUPHONICS
               sStr = "MM_EUPHONICS"
          Case MM_PRECEPT
               sStr = "MM_PRECEPT"
          Case MM_CRYSTAL_NET
               sStr = "MM_CRYSTAL_NET"
          Case MM_CHROMATIC
               sStr = "MM_CHROMATIC"
          Case MM_VOICEINFO
               sStr = "MM_VOICEINFO"
          Case MM_VIENNASYS
               sStr = "MM_VIENNASYS"
          Case MM_CONNECTIX
               sStr = "MM_CONNECTIX"
          Case MM_GADGETLABS
               sStr = "MM_GADGETLABS"
          Case MM_FRONTIER
               sStr = "MM_FRONTIER"
          Case MM_VIONA
               sStr = "MM_VIONA"
          Case MM_CASIO
               sStr = "MM_CASIO"
          Case MM_DIAMONDMM
               sStr = "MM_DIAMONDMM"
          Case MM_S3
               sStr = "MM_S3"
          Case MM_FRAUNHOFER_IIS
               sStr = "MM_FRAUNHOFER_IIS"
          Case Else
               sStr = "Unknown Manufacturer"
     End Select
     ' Return manufacturer
     GetManufacturer = sStr
End Function ' GetManufacturer

