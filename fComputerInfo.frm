VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fComputerInfo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Your Computer Components "
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   Icon            =   "fComputerInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picWait 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   2850
      ScaleHeight     =   1635
      ScaleWidth      =   4275
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2310
      Visible         =   0   'False
      Width           =   4275
      Begin MSComctlLib.ProgressBar prbDetecting 
         Height          =   315
         Left            =   165
         TabIndex        =   11
         Top             =   975
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Max             =   1
         Scrolling       =   1
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   1
         Left            =   90
         Picture         =   "fComputerInfo.frx":08CA
         Top             =   225
         Width           =   480
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "(please wait; this may take a moment)"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Index           =   3
         Left            =   885
         TabIndex        =   13
         Top             =   1305
         Width           =   2415
      End
      Begin VB.Label lb 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Identifying your computer components..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   210
         Index           =   2
         Left            =   705
         TabIndex        =   10
         Top             =   375
         Width           =   3405
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   1635
         Left            =   0
         Shape           =   4  'Gerundetes Rechteck
         Top             =   0
         Width           =   4275
      End
   End
   Begin VB.Frame fraDevices 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5310
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   4200
      Begin VB.CommandButton cmdExpCol 
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   675
         Width           =   1785
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh List"
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   270
         Width           =   1785
      End
      Begin MSComctlLib.TreeView trvComputer 
         Height          =   3870
         Left            =   120
         TabIndex        =   3
         Top             =   1290
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   6826
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   212
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
         MousePointer    =   99
         MouseIcon       =   "fComputerInfo.frx":1194
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   0
         Left            =   270
         Picture         =   "fComputerInfo.frx":14AE
         Top             =   345
         Width           =   480
      End
      Begin VB.Label lb 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Components"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   195
         TabIndex        =   17
         Top             =   1005
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   345
      Left            =   8700
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5550
      Width           =   1155
   End
   Begin VB.Frame fraProperties 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5310
      Left            =   4500
      TabIndex        =   6
      Top             =   150
      Width           =   5340
      Begin VB.CommandButton cmdAdjCols 
         Caption         =   "<-- Autosize &Columns -->"
         Height          =   285
         Left            =   3300
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   930
         Width           =   1920
      End
      Begin MSComctlLib.ListView lsvProperties 
         Height          =   3900
         Left            =   120
         TabIndex        =   5
         Top             =   1275
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   6879
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lb 
         Alignment       =   1  'Rechts
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   4
         Left            =   750
         TabIndex        =   16
         Top             =   630
         Width           =   420
      End
      Begin VB.Label lblComponentName 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1260
         TabIndex        =   15
         Top             =   585
         Width           =   3930
      End
      Begin VB.Label lblComponentType 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fest Einfach
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1260
         TabIndex        =   12
         Top             =   255
         Width           =   3930
      End
      Begin VB.Label lb 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Properties"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   210
         TabIndex        =   8
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label lb 
         Alignment       =   1  'Rechts
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Component Type"
         Height          =   390
         Index           =   0
         Left            =   300
         TabIndex        =   7
         Top             =   135
         Width           =   855
      End
   End
End
Attribute VB_Name = "fComputerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Uses WMI - Windows Management Instrumentation and tells you everything about your hardware that Windows knows

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LVM_SETCOLUMNWIDTH        As Long = &H1000 + 30
Private Const LVSCW_AUTOSIZE_USEHEADER  As Long = -2

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE     As Long = -16
Private Const WS_VSCROLL    As Long = &H200000

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sep1                As String
Private Sep2                As String
Private EffectiveWidth      As Long
Private SavedTop            As Long

Private Const Expand        As String = "Expand &all"
Private Const Collapse      As String = "Collapse &all"
Private Const Pfx           As String = "Win32_" 'prefix

Private Sub cmdAdjCols_Click()

  Dim ColumnHeader  As ColumnHeader
  Dim i             As Long
  Dim Wid           As Long

    With lsvProperties
        .Visible = False
        If .ListItems.Count Then
            If HasVScrollBar(lsvProperties) Then
                Wid = EffectiveWidth - 255
              Else 'HASVSCROLLBAR(LSVPROPERTIES) = FALSE/0
                Wid = EffectiveWidth
            End If
            For Each ColumnHeader In .ColumnHeaders
                SendMessage .hWnd, LVM_SETCOLUMNWIDTH, i, LVSCW_AUTOSIZE_USEHEADER
                i = i + 1
            Next ColumnHeader
          Else '.LISTITEMS.COUNT = FALSE/0
            .ColumnHeaders(1).Width = EffectiveWidth / 2
        End If
        DoEvents
        If .ColumnHeaders(1).Width + .ColumnHeaders(2).Width < Wid Then
            .ColumnHeaders(2).Width = Wid - .ColumnHeaders(1).Width
        End If
        .Visible = True
    End With 'LSVPROPERTIES
    On Error Resume Next
        trvComputer.SetFocus
    On Error GoTo 0

End Sub

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub cmdExpCol_Click()

  'expand collapse

  Dim Node  As Node

    With trvComputer
        .Visible = False
        If cmdExpCol.Caption = Expand Then
            For Each Node In .Nodes
                Node.Expanded = True
            Next Node
            cmdExpCol.Caption = Collapse
          Else 'NOT CMDEXPCOL.CAPTION...
            For Each Node In .Nodes
                Node.Expanded = (Node.Tag <> "C")
            Next Node
            ResetProperties
            cmdExpCol.Caption = Expand
        End If
        .Visible = True
        .SetFocus
        .Nodes(1).EnsureVisible
    End With 'TRVCOMPUTER

End Sub

Private Sub cmdRefresh_Click()

  'get Component classes and Components

  Dim Computer          As Variant
  Dim ComponentType     As Variant
  Dim Component         As Variant
  Dim ndSystem          As Node
  Dim ndComputer        As Node
  Dim ndComponentType   As Node
  Dim JustAdded         As String
  Dim i                 As Long

    cmdExit.Enabled = False
    cmdExpCol.Caption = Expand
    picWait.Visible = True
    For i = -picWait.Height To SavedTop
        picWait.Top = i
        Refresh
    Next i
    picWait.SetFocus
    Screen.MousePointer = vbHourglass
    Enabled = False
    trvComputer.Nodes.Clear
    ResetProperties
    DoEvents
    Set ndSystem = trvComputer.Nodes.Add(, , , "ComputerSystems")
    SetNodeProps ndSystem, True, True, vbBlack, True
    For Each Computer In GetComponents(Pfx & "ComputerSystem")
        Set ndComputer = trvComputer.Nodes.Add(ndSystem, tvwChild, , Computer)
        SetNodeProps ndComputer, True, True, &H6000&, True
        prbDetecting.Max = GetComponentTypes.Count  'adjust progress bar
        prbDetecting.Value = 0
        For Each ComponentType In GetComponentTypes
            If UBound(GetComponents(CStr(ComponentType))) >= 0 Then
                Set ndComponentType = trvComputer.Nodes.Add(ndComputer, tvwChild, , Replace$(ComponentType, Pfx, vbNullString))
                SetNodeProps ndComponentType, True, False, lblComponentType.ForeColor, True, "C"
                JustAdded = vbNullString
                For Each Component In GetComponents(CStr(ComponentType))
                    If Component = vbNullString Then
                        Component = "[No Info found]"
                    End If
                    If Component <> JustAdded Then 'not a duplicate
                        SetNodeProps trvComputer.Nodes.Add(ndComponentType, tvwChild, , Component), False, False, lblComponentName.ForeColor
                        JustAdded = Component
                    End If
                Next Component
                trvComputer.Refresh
            End If
            prbDetecting.Value = prbDetecting.Value + 1
    Next ComponentType, Computer
    ndSystem.EnsureVisible
    DoEvents
    Screen.MousePointer = vbDefault
    Sleep 400
    For i = picWait.Top To Height
        picWait.Top = i
        Refresh
    Next i
    picWait.Visible = False
    Enabled = True
    trvComputer.SetFocus
    cmdExit.Enabled = True

End Sub

Private Sub Form_Load()

    With lsvProperties
        .ColumnHeaders.Add , , "Property"
        .ColumnHeaders.Add , , "Value"
        .View = lvwReport
        EffectiveWidth = .Width - 75
    End With 'LSVPROPERTIES
    SavedTop = picWait.Top
    cmdExpCol.Caption = Expand
    Sep1 = Chr$(1) 'to be sure they do not appear in any national alphabet
    Sep2 = Chr$(2)
    Show
    DoEvents
    Sleep 400
    cmdRefresh_Click

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = Not cmdExit.Enabled

End Sub

Private Function GetComponents(ComponentType As String) As Variant

  'what it says - get all Components of a Component type

  Dim Components    As SWbemObjectSet
  Dim Component     As SWbemObject
  Dim sTemp         As String

    Set Components = GetObject("winmgmts:").InstancesOf(ComponentType)
    On Error Resume Next
        For Each Component In Components
            sTemp = sTemp & Component.Caption & Sep1
        Next Component
    On Error GoTo 0
    If Len(sTemp) Then
        sTemp = Left$(sTemp, Len(sTemp) - Len(Sep1))
    End If
    GetComponents = Split(sTemp, Sep1)

End Function

Private Function GetComponentTypes() As Collection

  'from http://msdn2.microsoft.com/en-us/library/aa389273.aspx

    Set GetComponentTypes = New Collection

    With GetComponentTypes
        .Add Pfx & "ComputerSystem"
        .Add Pfx & "Processor" 'represents a Component capable of interpreting a sequence of machine instructions on a Windows computer system.
        .Add Pfx & "DesktopMonitor" 'represents the type of monitor or display Component attached to the computer system.
        .Add Pfx & "DisplayConfiguration" 'represents configuration information for the display Component on a Windows system. This class is obsolete. In place of this class, use the properties in the .add Pfx & "VideoController, .add Pfx & "DesktopMonitor, and CIM_VideoControllerResolution classes.
        .Add Pfx & "DisplayControllerConfiguration" 'represents the video adapter configuration information of a Windows system. This class is obsolete. In place of this class, use the properties in the .add Pfx & "VideoController, .add Pfx & "DesktopMonitor, and CIM_VideoControllerResolution classes.
        .Add Pfx & "VideoController" 'represents the capabilities and management capacity of the video controller on a Windows computer system.
        .Add Pfx & "VideoSettings" 'relates a video controller and video settings that can be applied to it.
        .Add Pfx & "SoundComponent" 'represents the properties of a sound Component on a Windows computer system.
        .Add Pfx & "OnBoardComponent" 'represents common adapter Components built into the motherboard (system board).
        .Add Pfx & "Keyboard" 'represents a keyboard installed on a Windows system.
        .Add Pfx & "PointingComponent" 'represents an input Component used to point to and select regions on the display of a Windows computer system.
        .Add Pfx & "FloppyDrive" 'represents the capabilities of a floppy disk drive.
        .Add Pfx & "DiskDrive" 'represents a physical disk drive as seen by a computer running the Windows operating system.
        .Add Pfx & "AutochkSetting" 'represents the settings for the autocheck operation of a disk.
        .Add Pfx & "CDROMDrive" 'represents a CD-ROM drive on a Windows computer system.
        .Add Pfx & "TapeDrive" 'represents a tape drive on a Windows computer.
        .Add Pfx & "PhysicalMedia" 'represents any type of documentation or storage medium.
        .Add Pfx & "Printer" 'represents a Component connected to a Windows computer system that is capable of reproducing a visual image on a medium.
        .Add Pfx & "PrinterConfiguration" 'defines the configuration for a printer Component.
        .Add Pfx & "PrinterController" 'relates a printer and the local Component to which the printer is connected.
        .Add Pfx & "PrinterDriver" 'represents the drivers for a Win32_Printer instance."
        .Add Pfx & "PrinterDriverDll" 'relates a local printer and its driver file (not the driver itself).
        .Add Pfx & "PrinterSetting" 'relates a printer and its configuration settings.
        .Add Pfx & "PrintJob" 'represents a print job generated by a Windows application.
        .Add Pfx & "DriverForComponent" 'relates a printer to a printer driver.
        .Add Pfx & "TCPIPPrinterPort" 'represents a TCP/IP service access point.
        .Add Pfx & "MotherboardComponent" 'represents a Component that contains the central components of the Windows computer system.
        .Add Pfx & "BaseBoard" 'represents a baseboard (also known as a motherboard or system board).
        .Add Pfx & "BIOS" 'represents the attributes of the computer system's basic input/output services (BIOS) that are installed on the computer.
        .Add Pfx & "SMBIOSMemory" 'represents the capabilities and management of memory-related logical Components.
        .Add Pfx & "SystemBIOS" 'relates a computer system (including data such as startup properties, time zones, boot configurations, or administrative passwords) and a system BIOS"'services, languages, system management properties.
        .Add Pfx & "Bus" 'represents a physical bus as seen by a Windows operating system.
        .Add Pfx & "ComponentBus" 'relates a system bus and a logical Component using the bus.
        .Add Pfx & "ComponentSettings" 'relates a logical Component and a setting that can be applied to it.
        .Add Pfx & "PhysicalMemory" 'represents a physical memory Component located on a computer as available to the operating system.
        .Add Pfx & "PhysicalMemoryArray" 'represents details about the computer system's physical memory.
        .Add Pfx & "PhysicalMemoryLocation" 'relates an array of physical memory and its physical memory.
        .Add Pfx & "MemoryArray" 'represents the properties of the computer system memory array and mapped addresses.
        .Add Pfx & "MemoryArrayLocation" 'relates a logical memory array and the physical memory array upon which it exists.
        .Add Pfx & "MemoryComponent" 'represents the properties of a computer system's memory Component along with it's associated mapped addresses.
        .Add Pfx & "MemoryComponentArray" 'relates a memory Component and the memory array in which it resides.
        .Add Pfx & "MemoryComponentLocation" 'association class that relates a memory Component and the physical memory on which it exists.
        .Add Pfx & "ComponentMemoryAddress" 'represents a Component memory address on a Windows system.
        .Add Pfx & "CacheMemory" 'represents cache memory (internal and external) on a computer system.
        .Add Pfx & "FloppyController" 'represents the capabilities and management capacity of a floppy disk drive controller.
        .Add Pfx & "IDEController" 'represents the capabilities of an Integrated Drive Electronics (IDE) controller Component.
        .Add Pfx & "IDEControllerComponent" 'association class that relates an IDE controller and the logical Component.
        .Add Pfx & "SCSIController" 'represents a small computer system interface (SCSI) controller on a Windows system.
        .Add Pfx & "SCSIControllerComponent" 'relates a SCSI controller and the logical Component (disk drive) connected to it.
        .Add Pfx & "SystemSlot" 'represents physical connection points including ports, motherboard slots and peripherals, and proprietary connections points.
        .Add Pfx & "DMAChannel" 'represents a direct memory access (DMA) channel on a Windows computer system.
        .Add Pfx & "ParallelPort" 'represents the properties of a parallel port on a Windows computer system.
        .Add Pfx & "SerialPort" 'represents a serial port on a Windows system.
        .Add Pfx & "SerialPortConfiguration" 'represents the settings for data transmission on a Windows serial port.
        .Add Pfx & "SerialPortSetting" 'elates a serial port and its configuration settings.
        .Add Pfx & "InfraredComponent" 'represents the capabilities and management of an infrared Component.
        .Add Pfx & "NetworkAdapter" 'represents a network adapter on a Windows system.
        .Add Pfx & "NetworkAdapterConfiguration" 'represents the attributes and behaviors of a network adapter. The class is not guaranteed to be supported after the ratification of the Distributed"'management Task Force (DMTF) CIM network specification.
        .Add Pfx & "NetworkAdapterSetting" 'relates a network adapter and its configuration settings.
        .Add Pfx & "USBController" 'manages the capabilities of a universal serial bus (USB) controller.
        .Add Pfx & "USBControllerComponent" 'relates a USB controller and the CIM_LogicalComponent instances connected to it.
        .Add Pfx & "USBHub" 'represents the management characteristics of a USB hub.
        .Add Pfx & "ControllerHasHub" 'represents the hubs downstream from the universal serial bus (USB) controller.
        .Add Pfx & "PortConnector" 'represents physical connection ports, such as DB-25 pin male, Centronics, and PS/2.
        .Add Pfx & "PortResource" 'represents an I/O port on a Windows computer system.
        .Add Pfx & "SystemMemoryResource" 'represents a system memory resource on a Windows system.
        .Add Pfx & "AssociatedProcessorMemory" 'relates a processor and its cache memory.
        .Add Pfx & "PCMCIAController" 'manages the capabilities of a Personal Computer Memory Card Interface Adapter (PCMCIA) controller Component.
        .Add Pfx & "PNPEntity" 'represents the properties of a Plug and Play Component.
        .Add Pfx & "PNPComponent" 'relates a Component (known to Configuration Manager as a PNPEntity), and the function it performs.
        .Add Pfx & "PNPAllocatedResource" 'represents an association between logical Components and system resources.
        .Add Pfx & "SystemDriverPNPEntity" 'relates a Plug and Play Component on the Windows computer system and the driver that supports the Plug and Play Component.
        .Add Pfx & "POTSModem" 'represents the services and characteristics of a Plain Old Telephone Service (POTS) modem on a Windows system.
        .Add Pfx & "POTSModemToSerialPort" 'relates a modem and the serial port the modem uses.
        .Add Pfx & "1394Controller" 'represents the capabilities and management of a 1394 controller.
        .Add Pfx & "1394ControllerComponent" 'relates the high-speed serial bus (IEEE 1394 Firewire) Controller and the CIM_LogicalComponent instance connected to it.
        .Add Pfx & "AllocatedResource" 'relates a logical Component to a system resource.
        .Add Pfx & "IRQResource" 'represents an interrupt request line (IRQ) number on a Windows computer system.
        .Add Pfx & "Battery" 'represents a battery connected to the computer system.
        .Add Pfx & "PortableBattery" 'represents the properties of a portable battery, such as one used for a notebook computer.
        .Add Pfx & "AssociatedBattery" 'relates a logical Component and the battery it is using.
        .Add Pfx & "PowerManagementEvent" 'represents power management events resulting from power state changes.
        .Add Pfx & "UninterruptiblePowerSupply" 'represents the capabilities and management capacity of an uninterruptible power supply (UPS).
        .Add Pfx & "SystemEnclosure" 'represents the properties associated with a physical system enclosure.
        .Add Pfx & "Fan" 'represents the properties of a fan Component in the computer system.
        .Add Pfx & "HeatPipe" 'represents the properties of a heat pipe cooling Component.
        .Add Pfx & "Refrigeration" 'represents the properties of a refrigeration Component.
        .Add Pfx & "TemperatureProbe" 'represents the properties of a temperature sensor (electronic thermometer).
        .Add Pfx & "VoltageProbe" 'represents the properties of a voltage sensor (electronic voltmeter).
        .Add Pfx & "CurrentProbe" 'represents the properties of a current monitoring sensor (ammeter).
    End With 'GETComponentTYPES

End Function

Private Function GetProperties(ComponentId As Variant) As Variant

  Dim Components    As SWbemObjectSet
  Dim Component     As SWbemObject
  Dim Property      As SWbemProperty
  Dim sTemp         As String

    Set Components = GetObject("winmgmts:").InstancesOf(Pfx & ComponentId(2))
    For Each Component In Components 'search this particular Component
        If Component.Caption = ComponentId(3) Then
            For Each Property In Component.Properties_
                On Error Resume Next
                    sTemp = sTemp & Property.Name & Sep2 & Property.Value & Sep1
                On Error GoTo 0
            Next Property
            Exit For 'loop varying Component
        End If
    Next Component
    If Len(sTemp) Then
        sTemp = Left$(sTemp, Len(sTemp) - 1) 'remove the final sep1
    End If
    GetProperties = Split(sTemp, Sep1)

End Function

Private Function HasVScrollBar(Control As Control) As Boolean

    DoEvents
    HasVScrollBar = GetWindowLong(Control.hWnd, GWL_STYLE) And WS_VSCROLL

End Function

Private Sub lsvProperties_GotFocus()

    trvComputer.SetFocus

End Sub

Private Sub ResetProperties()

    With lsvProperties
        .ListItems.Clear
        .ColumnHeaders(1).Width = EffectiveWidth / 2
        .ColumnHeaders(2).Width = EffectiveWidth / 2
    End With 'LSVPROPERTIES
    lblComponentType = vbNullString
    lblComponentName = vbNullString

End Sub

Private Sub SetNodeProps(Node As Node, ByVal Bold As Boolean, ByVal Expanded As Boolean, ByVal ForeColor As Long, Optional ByVal EnsureVisible As Boolean = False, Optional Tag As String = vbNullString)

    With Node
        .Bold = Bold
        .Expanded = Expanded
        .ForeColor = ForeColor
        .Tag = Tag
        If EnsureVisible Then
            .EnsureVisible
        End If
    End With 'NODE

End Sub

Private Sub trvComputer_NodeClick(ByVal Node As MSComctlLib.Node)

  'properties and values to listbox

  Dim arrFullPath   As Variant
  Dim arrItems      As Variant
  Dim Property      As Variant

    arrFullPath = Split(Node.FullPath, "\")
    If UBound(arrFullPath) = 3 Then 'user clicked a Component
        Enabled = False
        Screen.MousePointer = vbHourglass
        ResetProperties
        lblComponentType = " " & arrFullPath(2)
        lblComponentName = " " & arrFullPath(3)
        For Each Property In GetProperties(arrFullPath)
            arrItems = Split(Replace$(Property, vbTab, " "), Sep2)
            If UBound(arrItems) <> 1 Then
                arrItems = Split("[No Property found]/[No Info available]", "/")
            End If
            If Len(arrItems(1)) Then
                If arrItems(0) <> "SystemName" Then 'to suppress systemname
                    lsvProperties.ListItems.Add(, , CStr(arrItems(0))).SubItems(1) = arrItems(1)
                End If
            End If
        Next Property
        cmdAdjCols_Click
        Screen.MousePointer = vbDefault
        Enabled = True
      Else 'NOT UBOUND(ARRFULLPATH)...
        ResetProperties
    End If

End Sub

':) Ulli's VB Code Formatter V2.22.15 (2007-Feb-13 12:26)  Decl: 22  Code: 380  Total: 402 Lines
':) CommentOnly: 6 (1,5%)  Commented: 107 (26,6%)  Empty: 60 (14,9%)  Max Logic Depth: 6
