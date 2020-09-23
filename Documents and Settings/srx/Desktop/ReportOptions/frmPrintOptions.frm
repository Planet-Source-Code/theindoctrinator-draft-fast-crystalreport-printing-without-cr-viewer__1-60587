VERSION 5.00
Begin VB.Form frmPrintOptions 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DFBF9F&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FACOR ERP : Printer Options"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00DFBF9F&
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   4965
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2535
      Width           =   1380
   End
   Begin VB.ListBox lstManufacturer 
      BackColor       =   &H00DFBF9F&
      Height          =   1845
      IntegralHeight  =   0   'False
      Left            =   75
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   345
      Width           =   2085
   End
   Begin VB.ListBox lstPrinters 
      BackColor       =   &H00DFBF9F&
      Height          =   1845
      IntegralHeight  =   0   'False
      Left            =   2235
      TabIndex        =   1
      Top             =   345
      Width           =   4080
   End
   Begin VB.CommandButton cmdInstall 
      BackColor       =   &H00DFBF9F&
      Caption         =   "&Install"
      Enabled         =   0   'False
      Height          =   360
      Left            =   3405
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2535
      Width           =   1380
   End
   Begin VB.CommandButton cmdReset 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFBF9F&
      Caption         =   "&Reset"
      Height          =   360
      Left            =   1860
      MaskColor       =   &H00DFBF9F&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2535
      Width           =   1380
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   6945
      Y1              =   2355
      Y2              =   2355
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6945
      Y1              =   2370
      Y2              =   2370
   End
   Begin VB.Label lblManufacturers 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFBF9F&
      BackStyle       =   0  'Transparent
      Caption         =   "&Manufacturers:"
      Height          =   195
      Left            =   75
      TabIndex        =   6
      Top             =   105
      Width           =   1065
   End
   Begin VB.Label lblPrinters 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFBF9F&
      BackStyle       =   0  'Transparent
      Caption         =   "&Printers:"
      Height          =   195
      Left            =   2235
      TabIndex        =   3
      Top             =   105
      Width           =   570
   End
End
Attribute VB_Name = "frmPrintOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim objRPT As New clsPrinter
Dim INF As New clsPrinter
Private Sub EnumManufacturers()
Dim Manufacturers() As String
Dim i As Integer
    
    INF.EnumSectionKeys "manufacturer", Manufacturers()
    
    lstManufacturer.Clear
    
    For i = 0 To UBound(Manufacturers)
        lstManufacturer.AddItem Manufacturers(i)
    Next

End Sub

'Private Sub cmdHaveDisk_Click()
'
'    On Error Resume Next
'
'    With CommonDialog1
'        .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNFileMustExist
'        .Filter = "Setup Information files|*.inf"
'        .ShowOpen
'    End With
'
'    If Err.Number = 0 Then
'
'        INF.OpenInf CommonDialog1.FileName, "printer"
'
'        EnumManufacturers
'
'    End If
'
'End Sub

Private Sub cmdCancel_Click()
'    MsgBox GetOSVersion
    Unload Me
End Sub

Private Sub cmdInstall_Click()
Dim ManufacturerSection As String, InstallSection As String
Dim DriverFile As String, PrinterName As String
Dim DataFile As String, DataSection As String, ConfigFile As String
     
    ManufacturerSection = INF.GetValue("Manufacturer", lstManufacturer.List(lstManufacturer.ListIndex))
    PrinterName = lstPrinters.List(lstPrinters.ListIndex)
    
    InstallSection = INF.GetField(ManufacturerSection, PrinterName, 1)
    
    ' Execute section. InstallSection returns true
    ' if the files are installed.
    
    If INF.InstallSection(InstallSection, , , Me.hwnd) Then
        
        ' Get DataSection
        DataSection = INF.GetValue(InstallSection, "DataSection")
        
        ' Get the driver file
        DriverFile = INF.GetValue(InstallSection, "DriverFile")
        
        ' If there's no DriverFile key in the
        ' section, look at DataSection
        If DriverFile = "" Then DriverFile = INF.GetValue(DataSection, "DriverFile")
        
        ' If the driver file is still empty
        ' use the InstallSection name
        ' as the driver file
        If DriverFile = "" Then DriverFile = InstallSection
        
        ' Get the data file
        DataFile = INF.GetValue(InstallSection, "DataFile")
        If DataFile = "" Then DataFile = INF.GetValue(DataSection, "DataFile")
        If DataFile = "" Then DataFile = InstallSection
                
        ConfigFile = INF.GetValue(DataSection, "ConfigFile")
        If ConfigFile = "" Then ConfigFile = DriverFile
        
        ' Add the printer driver
        ' to windows
        AddNewDriver PrinterName, DriverFile, DataFile, ConfigFile
        
        ' Add a new printer with the
        ' installed printer driver
        AddNewPrinter PrinterName, PrinterName
        
    End If
    DoEvents
    Unload Me
End Sub

Private Sub cmdReset_Click()
        
    Set INF = New clsPrinter
    
    
   Dim osinfo As OSVERSIONINFO
   Dim retvalue As Integer
   osinfo.dwOSVersionInfoSize = 148
   osinfo.szCSDVersion = Space$(128)
   retvalue = GetVersionExA(osinfo)
   'Specify the DSN parameters.
    If osinfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
        'this is Win 98/95
        INF.OpenInf "msprint.inf", "printer"
    Else
        'this is Win NT+/2k/XP
        INF.OpenInf "ntprint.inf", "printer"
    End If
    
    INF.AppendInf "msprint2.inf"
    INF.AppendInf "layout.inf"
    
    EnumManufacturers

End Sub

Private Sub Form_Activate()
        Dim strMaufac As String
        Dim strPrinter As String
        Screen.MousePointer = vbDefault
        strMaufac = GetStringFromList("Generic", lstManufacturer)
        SetStringInList strMaufac, lstManufacturer
        DoEvents
        'Call lstManufacturer_Click
        'DoEvents
        strPrinter = GetStringFromList("Generic / Text Only", lstPrinters)
        SetStringInList strPrinter, lstPrinters
                
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    cmdReset_Click
End Sub


Private Sub lstManufacturer_Click()
Dim ManufSection As String
Dim Prntrs() As String, i As Integer

    ' Get the manufacturer's section.
    ManufSection = INF.GetValue("Manufacturer", lstManufacturer.List(lstManufacturer.ListIndex))
    
    ' Enumaret printer from
    ' selected manufacturer
    INF.EnumSectionKeys ManufSection, Prntrs()
    
    ' Add printers to listbox
    
    lstPrinters.Clear
    
    For i = 0 To UBound(Prntrs)
        lstPrinters.AddItem Prntrs(i)
    Next
    
End Sub

Private Sub lstPrinters_Click()

    cmdInstall.Enabled = True
    
End Sub



