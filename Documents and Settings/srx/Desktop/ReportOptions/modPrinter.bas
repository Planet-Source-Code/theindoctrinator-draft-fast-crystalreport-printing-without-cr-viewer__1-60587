Attribute VB_Name = "modPrinter"
'GLOBAL Varialble Declaration For Printer
Global CommandLineArgsSent As String
Global PreviousPrinterDevName As String
Global PreviousPrinterDevPort As String
Global PreviousprinterDevDriver As String
Global PreviousDefaultPrinter As String
Global Const SetDefaultPrinterTo As String = "Generic / Text Only"
Global CurrentDefaultPrinter As String
Global OpenReportForEmail   As Boolean
Global prt As Printer  'this value will hold the DefaultPrinter for the CurrentApplication, gets set in the EmailModule
Global Rprt As Printer 'this value will hold the Printer for a given report, if the report is set to use a specific printer we need to know what the printer was before we reset it.
Global PRINTER_NAME As String
Public PrinterFormNames As fFromNames

Public Const HWND_BROADCAST = &HFFFF
'Private Const HWND_BROADCAST As Long = 65535
Public Const WM_WININICHANGE = &H1A
'Private Const WM_WININICHANGE As Integer = 26


'GetPrinterDriver Path
Public Const MAX_PATH = 260


' constants for DEVMODE structure
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32
Public Const DM_FORMNAME As Long = &H10000
Public Const DM_ORIENTATION = &H1&
Public Const DM_PAPERSIZE = &H2&
Public Const DM_PAPERLENGTH = &H4&
Public Const DM_PAPERWIDTH = &H8&
' Constants for DocumentProperties() call
Public Const DM_MODIFY = 8
Public Const DM_IN_BUFFER = DM_MODIFY
Public Const DM_COPY = 2
Public Const DM_OUT_BUFFER = DM_COPY

' constants for DesiredAccess member of PRINTER_DEFAULTS
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8

Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)
' constant that goes into PRINTER_INFO_5 Attributes member
' to set it as default
Public Const PRINTER_ATTRIBUTE_DEFAULT = 4
'Constant for OSVERSIONINFO.dwPlatformId
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
'The WM_WININICHANGE message is obsolete. It is included for
'compatibility with earlier versions of the system. New
'applications should use the WM_SETTINGCHANGE message.

' Custom constants for this sample's SelectForm function
Public Const FORM_NOT_SELECTED = 0
Public Const FORM_SELECTED = 1
Public Const FORM_ADDED = 2


Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

 Public Const PD_ALLPAGES = &H0
   Public Const PD_COLLATE = &H10
   Public Const PD_DISABLEPRINTTOFILE = &H80000
   Public Const PD_ENABLEPRINTHOOK = &H1000
   Public Const PD_ENABLEPRINTTEMPLATE = &H4000
   Public Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
   Public Const PD_ENABLESETUPHOOK = &H2000
   Public Const PD_ENABLESETUPTEMPLATE = &H8000
   Public Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
   Public Const PD_HIDEPRINTTOFILE = &H100000
   Public Const PD_NONETWORKBUTTON = &H200000
   Public Const PD_NOPAGENUMS = &H8
   Public Const PD_NOSELECTION = &H4
   Public Const PD_NOWARNING = &H80
   Public Const PD_PAGENUMS = &H2
   Public Const PD_PRINTSETUP = &H40
   Public Const PD_PRINTTOFILE = &H20
   Public Const PD_RETURNDC = &H100
   Public Const PD_RETURNDEFAULT = &H400
   Public Const PD_RETURNIC = &H200
   Public Const PD_SELECTION = &H1
   Public Const PD_SHOWHELP = &H800
   Public Const PD_USEDEVMODECOPIES = &H40000
   Public Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000

   ' Constants for PAGESETUPDLG
   Public Const PSD_DEFAULTMINMARGINS = &H0
   Public Const PSD_DISABLEMARGINS = &H10
   Public Const PSD_DISABLEORIENTATION = &H100
   Public Const PSD_DISABLEPAGEPAINTING = &H80000
   Public Const PSD_DISABLEPAPER = &H200
   Public Const PSD_DISABLEPRINTER = &H20
   Public Const PSD_ENABLEPAGEPAINTHOOK = &H40000
   Public Const PSD_ENABLEPAGESETUPHOOK = &H2000
   Public Const PSD_ENABLEPAGESETUPTEMPLATE = &H8000
   Public Const PSD_ENABLEPAGESETUPTEMPLATEHANDLE = &H20000
   Public Const PSD_INHUNDREDTHSOFMILLIMETERS = &H8
   Public Const PSD_INTHOUSANDTHSOFINCHES = &H4
   Public Const PSD_INWININIINTLMEASURE = &H0
   Public Const PSD_MARGINS = &H2
   Public Const PSD_MINMARGINS = &H1
   Public Const PSD_NOWARNING = &H80
   Public Const PSD_RETURNDEFAULT = &H400
   Public Const PSD_SHOWHELP = &H800

   ' Custom Global Constants
   Public Const DLG_PRINT = 0
   Public Const DLG_PRINTSETUP = 1

Public Type RECTL
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type SIZEL
        cx As Long
        cy As Long
End Type

Public Type SECURITY_DESCRIPTOR
        Revision As Byte
        Sbz1 As Byte
        Control As Long
        Owner As Long
        Group As Long
        Sacl As Long  ' ACL
        Dacl As Long  ' ACL
End Type




Public INSTALLED_PRINTERS ' List of All Available Printers(Used AS Array)
'Public
Public Type PrinterInfo
    PrinterName As String
    PrinterPort As String
    PrinterDriver As String
    PrinterDevice As String
End Type

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

' The two definitions for FORM_INFO_1 make the coding easier.
Public Type FORM_INFO_1
        Flags As Long
        pName As Long   ' String
        Size As SIZEL
        ImageableArea As RECTL
End Type

Public Type sFORM_INFO_1
        Flags As Long
        pName As String
        Size As SIZEL
        ImageableArea As RECTL
End Type

Public Type DevMode
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmLogPixels As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    dmICMMethod As Long ' // Windows 95 only
    dmICMIntent As Long ' // Windows 95 only
    dmMediaType As Long ' // Windows 95 only
    dmDitherType As Long ' // Windows 95 only
    dmReserved1 As Long ' // Windows 95 only
    dmReserved2 As Long ' // Windows 95 only
End Type


'this is for the Apollo Windows NT/2000 driver version 2.4

Public Type PRIVATEDEVMODE
   dwPrivDATA As Integer                   ' private data id
   interpret As Long                               ' Interpretation yes/no bitmap
   security As Byte                                ' PDF 417 security
   qrerrlevel As Byte                              ' QR Error Correction Level
   unused As Byte                                 ' Open for future expansion
   blanksub As Integer              ' Blank substitute character
   pdfrowsize As Integer           ' Codewords/row PDF417 0 means calculate
   dmtrxrowsize As Integer                    ' Data Matrix code cell size in Dots
   qrrowsize As Integer              ' QR code cell size in Dots
   demandoffset As Integer                    ' Presentation position
   LeftMargin As Integer              ' Left offset
   labeloffset As Integer             ' Vertical offset
   lgapsize As Integer                             ' distance between labels
   heat As Byte                                      ' heat 0 based
   sensor As Byte                                  ' 0-see through sensor, 1-reflective sensor
   demandmode As Byte                       ' 0-batchmode,1-peel off,2-tear off
   directthermal As Byte           ' 0-thermal transfer 1-direct thermal;
   continuous As Byte               ' 0-die-cut; 1-continuous
   speed As Byte                                   ' 0 based speed ordinal
   flip As Byte                                       ' 1-180 degree flip
   mirror As Byte                                   ' 0-normal,1-mirror
   negative As Byte                                ' 0-normal,1-negative
   intrpstartstop As Byte           ' Include start stop in interpretation
   model As Byte                                   ' printer model code
   ribbonsave As Byte                           ' 0-do not use ribbon saver
   cut As Byte                                       ' 0-no cut,1-cut after label,2-cut after batch,3-cut after job
   cardload As Byte                               ' 0-normal,1-cardload
   formatname As String * 9                  ' Name of format when loading card
End Type



Public Type PRINTER_INFO_5
    pPrinterName As String
    pPortName As String
    Attributes As Long
    DeviceNotSelectedTimeout As Long
    TransmissionRetryTimeout As Long
End Type

Public Type PRINTER_DEFAULTS
    pDatatype As Long
    pDevMode As Long
    DesiredAccess As Long
End Type


Type PRINTER_INFO_2
    pServerName As String
    pPrinterName As String
    pShareName As String
    pPortName As String
    pDriverName As String
    pComment As String
    pLocation As String
    pDevMode As Long
    pSepFile As String
    pPrintProcessor As String
    pDatatype As String
    pParameters As String
    pSecurityDescriptor As Long
    Attributes As Long
    Priority As Long
    DefaultPriority As Long
    StartTime As Long
    UntilTime As Long
    Status As Long
    cJobs As Long
    AveragePPM As Long
End Type

'Public Type PRINTER_INFO_2
'        pServerName As String
'        pPrinterName As String
'        pShareName As String
'        pPortName As String
'        pDriverName As String
'        pComment As String
'        pLocation As String
'        pDevMode As DEVMODE
'        pSepFile As String
'        pPrintProcessor As String
'        pDatatype As String
'        pParameters As String
'        pSecurityDescriptor As SECURITY_DESCRIPTOR
'        Attributes As Long
'        Priority As Long
'        DefaultPriority As Long
'        StartTime As Long
'        UntilTime As Long
'        Status As Long
'        cJobs As Long
'        AveragePPM As Long
'End Type
'Type for Set Printer FormSize
Public Type PSDEVMODE

    dmPublic As DevMode             ' public portion

    dmPrivate As PRIVATEDEVMODE     ' private portion

End Type


Type INFCONTEXT
    INF As Long
    CurrentInf As Long
    Section As Long
    Line As Long
End Type

Type DRIVER_INFO_2
        cVersion As Long
        pName As String
        pEnvironment As String
        pDriverPath As String
        pDataFile As String
        pConfigFile As String
End Type
'Must Add A Value of Newly inserted Form
'Cautions :
'Use PreFix 'f' Before Values
'Use '_' instead of '.'
'EG:
' 10 X 12 To Be Entered as "f10X12"
' 8.5 X 12 To Be Entered as "f8_5X12"
'-------------- IMPORTANT ------------------
' Add the newly added Form in ValidateForms function
' This Function is in MDI STORE
Public Type fFromNames
    f14X12 As String
    f14X6 As String
    f10X12 As String
    f10X6 As String
    f8_5X12 As String
    f14X14 As String
    f14X15 As String
End Type


Type PRINTDLG_TYPE
           lStructSize As Long
           hwndOwner As Long
           hDevMode As Long
           hDevNames As Long
           hdc As Long
           Flags As Long
           nFromPage As Integer
           nToPage As Integer
           nMinPage As Integer
           nMaxPage As Integer
           nCopies As Integer
           hInstance As Long
           lCustData As Long
           lpfnPrintHook As Long
           lpfnSetupHook As Long
           lpPrintTemplateName As String
           lpSetupTemplateName As String
           hPrintTemplate As Long
           hSetupTemplate As Long
   End Type

   Type DEVNAMES_TYPE
           wDriverOffset As Integer
           wDeviceOffset As Integer
           wOutputOffset As Integer
           wDefault As Integer
           extra As String * 100
   End Type

   Type DEVMODE_TYPE
           dmDeviceName As String * CCHDEVICENAME
           dmSpecVersion As Integer
           dmDriverVersion As Integer
           dmSize As Integer
           dmDriverExtra As Integer
           dmFields As Long
           dmOrientation As Integer
           dmPaperSize As Integer
           dmPaperLength As Integer
           dmPaperWidth As Integer
           dmScale As Integer
           dmCopies As Integer
           dmDefaultSource As Integer
           dmPrintQuality As Integer
           dmColor As Integer
           dmDuplex As Integer
           dmYResolution As Integer
           dmTTOption As Integer
           dmCollate As Integer
           dmFormName As String * CCHFORMNAME
           dmUnusedPadding As Integer
           dmBitsPerPel As Integer
           dmPelsWidth As Long
           dmPelsHeight As Long
           dmDisplayFlags As Long
           dmDisplayFrequency As Long
   End Type


   ' type definitions:
   Public Type Rect
         Left As Long
         Top As Long
         Right As Long
         Bottom As Long
   End Type
   
   Public Type POINTAPI
           x As Long
           y As Long
   End Type
 Type PRINTSETUPDLG_TYPE
           lStructSize As Long
           hwndOwner As Long
           hDevMode As Long
           hDevNames As Long
           Flags As Long
           ptPaperSize As POINTAPI
           rtMinMargin As Rect
           rtMargin As Rect
           hInstance As Long
           lCustData As Long
           lpfnPageSetupHook As Long ' LPPAGESETUPHOOK
           lpfnPagePaintHook As Long ' LPPAGESETUPHOOK
           lpPageSetupTemplateName As String
           hPageSetupTemplate As Long ' HGLOBAL
   End Type


Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Public Declare Function EnumForms Lib "winspool.drv" Alias "EnumFormsA" (ByVal hPrinter As Long, ByVal Level As Long, ByRef pForm As Any, ByVal cbBuf As Long, ByRef pcbNeeded As Long, ByRef pcReturned As Long) As Long
Public Declare Function AddForm Lib "winspool.drv" Alias "AddFormA" (ByVal hPrinter As Long, ByVal Level As Long, pForm As Byte) As Long
Public Declare Function DeleteForm Lib "winspool.drv" Alias "DeleteFormA" (ByVal hPrinter As Long, ByVal pFormName As String) As Long
Public Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Public Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hdc As Long, lpInitData As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function AddPrinterDriver Lib "winspool.drv" Alias "AddPrinterDriverA" (ByVal pName As String, ByVal Level As Long, pDriverInfo As Any) As Long
'Public Declare Function AddPrinter Lib "winspool.drv" Alias "AddPrinterA" (ByVal pName As String, ByVal Level As Long, pPrinter As Any) As Long
Public Declare Function AddPrinter Lib "winspool.drv" Alias "AddPrinterA" (ByVal pName As String, ByVal Level As Long, pPrinter As Any) As Long
' Optional functions not used in this sample, but may be useful.
'Public Declare Function GetForm Lib "winspool.drv" Alias "GetFormA" (ByVal hPrinter As Long, ByVal pFormName As String, ByVal Level As Long, pForm As Byte, ByVal cbBuf As Long, pcbNeeded As Long) As Long
'Public Declare Function SetForm Lib "winspool.drv" Alias "SetFormA" (ByVal hPrinter As Long, ByVal pFormName As String, ByVal Level As Long, pForm As Byte) As Long

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Long) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long
Public Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Public Declare Function PrinterProperties Lib "winspool.drv" (ByVal hwnd As Long, ByVal hPrinter As Long) As Long
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Any) As Long

Public Declare Function GetLastError Lib "kernel32" () As Integer
Public Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Public Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long

Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINESCROLL = &HB6
'*********************************************************************************************
'
' .INF file parser class
'
' Setup API declarations
'
'*********************************************************************************************

'*********************************************************************************************

Declare Function SetupOpenInfFile Lib "setupapi" Alias "SetupOpenInfFileA" (ByVal FileName As String, ByVal InfClass As String, ByVal InfStyle As InfStyles, ErrorLine As Long) As Long
Declare Sub SetupCloseInfFile Lib "setupapi" (ByVal InfHandle As Long)
Declare Function SetupOpenAppendInfFile Lib "setupapi" Alias "SetupOpenAppendInfFileA" (ByVal FileName As String, ByVal InfHandle As Long, ErrorLine As Long) As Boolean
Declare Function SetupFindFirstLine Lib "setupapi" Alias "SetupFindFirstLineA" (ByVal hInf As Long, ByVal Section As String, ByVal Key As String, Context As INFCONTEXT) As Boolean
Declare Function SetupFindNextLine Lib "setupapi" Alias "SetupFindNextLineA" (ContextIn As INFCONTEXT, ContextOut As INFCONTEXT) As Boolean
Declare Function SetupFindNextMatchLine Lib "setupapi" Alias "SetupFindNextMatchLineA" (ContextIn As INFCONTEXT, ByVal Key As String, ContextOut As INFCONTEXT) As Boolean
Declare Function SetupGetLineByIndex Lib "setupapi" Alias "SetupGetLineByIndexA" (ByVal InfHandle As Long, ByVal Section As String, ByVal Index As Long, Context As INFCONTEXT) As Boolean
Declare Function SetupGetLineCount Lib "setupapi" Alias "SetupGetLineCountA" (ByVal InfHandle As Long, ByVal Section As String) As Long
Declare Function SetupGetLineText Lib "setupapi" Alias "SetupGetLineTextA" (Context As Any, ByVal InfHandle As Long, ByVal Section As String, ByVal Key As String, ByVal ReturnBuffer As String, ByVal ReturnBufferSize As Long, RequiredSize As Long) As Boolean
Declare Function SetupGetStringField Lib "setupapi" Alias "SetupGetStringFieldA" (Context As Any, ByVal FieldIndex As Long, ByVal ReturnBuffer As String, ByVal ReturnBufferSize As Long, RequiredSize As Long) As Boolean
Declare Function SetupInitDefaultQueueCallback Lib "setupapi" (ByVal OwnerWindow As Long) As Long
Declare Function SetupInstallFromInfSection Lib "setupapi" Alias "SetupInstallFromInfSectionA" (ByVal Owner As Long, ByVal InfHandle As Long, ByVal SectionName As String, ByVal Flags As InstallFlags, ByVal RelativeKeyRoot As Long, ByVal SourceRootPath As String, ByVal CopyFlags As CopyFlags, ByVal MsgHandler As Long, ByVal Context As Long, ByVal DeviceInfoSet As Long, ByVal DeviceInfoData As Any) As Boolean
Public Declare Function GetPrinterDriverDirectory Lib "winspool.drv" Alias "GetPrinterDriverDirectoryA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, ByVal pDriverDirectory As String, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function PageSetupDialog Lib "comdlg32.dll" Alias "PageSetupDlgA" (pSetupPrintdlg As PRINTSETUPDLG_TYPE) As Long
Public Declare Function SetForm Lib "winspool.drv" Alias "SetFormA" (ByVal hPrinter As Long, ByVal pFormName As String, ByVal Level As Long, pForm As Byte) As Long
'Enum InfStyles
'    INF_STYLE_NONE = 0       ' unrecognized or non-existent
'    INF_STYLE_OLDNT = 1      ' winnt 3.x
'    INF_STYLE_WIN4 = 2       ' Win95
'End Enum


Enum InstallFlags
    SPINST_LOGCONFIG = &H1
    SPINST_INIFILES = &H2
    SPINST_REGISTRY = &H4
    SPINST_INI2REG = &H8
    SPINST_FILES = &H10
    SPINST_ALL = &H1F
End Enum

Enum InfStyles
    INF_STYLE_NONE = 0       ' unrecognized or non-existent
    INF_STYLE_OLDNT = 1      ' winnt 3.x
    INF_STYLE_WIN4 = 2       ' Win95
End Enum

Public Enum CopyFlags
    SP_COPY_DELETESOURCE = &H1               ' delete source file on successful copy
    SP_COPY_REPLACEONLY = &H2                ' copy only if target file already present
    SP_COPY_NEWER = &H4                      ' copy only if source file newer than target
    SP_COPY_NOOVERWRITE = &H8                ' copy only if target doesn't exist
    SP_COPY_NODECOMP = &H10                  ' don't decompress source file while copying
    SP_COPY_LANGUAGEAWARE = &H20             ' don't overwrite file of different language
    SP_COPY_SOURCE_ABSOLUTE = &H40           ' SourceFile is a full source path
    SP_COPY_SOURCEPATH_ABSOLUTE = &H80       ' SourcePathRoot is the full path
    SP_COPY_IN_USE_NEEDS_REBOOT = &H100      ' System needs reboot if file in use
    SP_COPY_FORCE_IN_USE = &H200             ' Force target-in-use behavior
    SP_COPY_NOSKIP = &H400                   ' Skip is disallowed for this file or section
    SP_FLAG_CABINETCONTINUATION = &H800      ' Used with need media notification
    SP_COPY_FORCE_NOOVERWRITE = &H1000       ' like NOOVERWRITE but no callback nofitication
    SP_COPY_FORCE_NEWER = &H2000             ' like NEWER but no callback nofitication
    SP_COPY_WARNIFSKIP = &H4000              ' system critical file: warn if user tries to skip
    SP_COPY_NOBROWSE = &H8000                ' Browsing is disallowed for this file or section
End Enum


Public Function GetFormName(ByVal PrinterHandle As Long, _
                          FormSize As SIZEL, FormName As String) As Integer
    Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1           ' Working FI1 array
    Dim Temp() As Byte                  ' Temp FI1 array
    Dim FormIndex As Integer
    Dim FName As String
    Dim BytesNeeded As Long
    Dim RetVal As Long
    
    FormName = Replace(FormName, " ", "", 1, Len(FormName))
    FormName = UCase(Trim(FormName))
    FormIndex = 0
    ReDim aFI1(1)
    ' First call retrieves the BytesNeeded.
    RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
    ReDim Temp(BytesNeeded)
    ReDim aFI1(BytesNeeded / Len(FI1))
    ' Second call actually enumerates the supported forms.
    RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, _
             NumForms)
    Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
    For i = 0 To NumForms - 1
        With aFI1(i)
            FName = PtrCtoVbString(.pName)
            FName = Replace(FName, " ", "", 1, Len(FName))
            FName = UCase(Trim(FName))
'            If FName = FormName Then
'                MsgBox "Found"
'            End If
            If .Size.cx = FormSize.cx And .Size.cy = FormSize.cy And FormName = FName Then
               ' Found the desired form
                
                FormIndex = i + 1
                Exit For
            End If
        End With
    Next i
    GetFormName = FormIndex  ' Returns non-zero when form is found.
End Function

Public Function AddNewForm(PrinterHandle As Long, FormSize As SIZEL, _
                           FormName As String) As String
    Dim FI1 As sFORM_INFO_1
    Dim aFI1() As Byte
    Dim RetVal As Long
    
    With FI1
        .Flags = 0
        .pName = FormName
        With .Size
            .cx = FormSize.cx
            .cy = FormSize.cy
        End With
        With .ImageableArea
            .Left = 0
            .Top = 0
            .Right = FI1.Size.cx
            .Bottom = FI1.Size.cy
        End With
    End With
    ReDim aFI1(Len(FI1))
    Call CopyMemory(aFI1(0), FI1, Len(FI1))
    RetVal = AddForm(PrinterHandle, 1, aFI1(0))
    If RetVal = 0 Then
        If Err.LastDllError = 5 Then
            MsgBox "You do not have permissions to add a form to " & _
               Printer.DeviceName, vbExclamation, "Access Denied!"
        Else
            MsgBox "Error: " & Err.LastDllError, "Error Adding Form"
        End If
        AddNewForm = "none"
    Else
        AddNewForm = FI1.pName
    End If
End Function

Public Function SelectForm(FormName As String, ByVal MyhWnd As Long) As Integer
    Dim nSize As Long           ' Size of DEVMODE
    Dim pDevMode As DevMode
    Dim PrinterHandle As Long   ' Handle to printer
    Dim hPrtDC As Long          ' Handle to Printer DC
    Dim PrinterName As String
    Dim aDevMode() As Byte      ' Working DEVMODE
    Dim FormSize As SIZEL
    
    PrinterName = Printer.DeviceName  ' Current printer
    hPrtDC = Printer.hdc              ' hDC for current Printer
    SelectForm = FORM_NOT_SELECTED    ' Set for failure unless reset in code.
    
    ' Get a handle to the printer.
   If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ' Retrieve the size of the DEVMODE.
        nSize = DocumentProperties(MyhWnd, PrinterHandle, PrinterName, 0&, _
                0&, 0&)
        ' Reserve memory for the actual size of the DEVMODE.
        ReDim aDevMode(1 To nSize)
    
        ' Fill the DEVMODE from the printer.
        nSize = DocumentProperties(MyhWnd, PrinterHandle, PrinterName, _
                aDevMode(1), 0&, DM_OUT_BUFFER)
        ' Copy the Public (predefined) portion of the DEVMODE.
        Call CopyMemory(pDevMode, aDevMode(1), Len(pDevMode))
        
        
        If FormName = "MyCustomForm" Then

            ' Use form "MyCustomForm", adding it if necessary.
            ' Set the desired size of the form needed.
            With FormSize   ' Given in thousandths of millimeters
                .cx = 214000   ' width
                .cy = 216000   ' height
            End With
            
            If GetFormName(PrinterHandle, FormSize, FormName) = 0 Then
                ' Form not found - Either of the next 2 lines will work.
                'FormName = AddNewForm(PrinterHandle, FormSize, "MyCustomForm")
                AddNewForm PrinterHandle, FormSize, "MyCustomForm"
                If GetFormName(PrinterHandle, FormSize, FormName) = 0 Then
                    ClosePrinter (PrinterHandle)
                    SelectForm = FORM_NOT_SELECTED   ' Selection Failed!
                    Exit Function
                Else
                    SelectForm = FORM_ADDED  ' Form Added, Selection succeeded!

                End If
             Else

            End If
        End If
        
        Dim SZ
        SZ = Split(UCase(FormName), "X", Len(FormName))
        With FormSize   ' Given in thousandths of millimeters
            .cx = SZ(0) * 2.54 * 10 * 1000
            .cy = SZ(1) * 2.54 * 10 * 1000
        End With

        If GetFormName(PrinterHandle, FormSize, FormName) <> 0 Then
            pDevMode.dmPaperSize = GetFormName(PrinterHandle, FormSize, FormName)

        End If
        
        ' Change the appropriate member in the DevMode.
        ' In this case, you want to change the form name.
        pDevMode.dmFormName = FormName & Chr(0)  ' Must be NULL terminated!
        
        ' Set the dmFields bit flag to indicate what you are changing.
        pDevMode.dmFields = DM_FORMNAME Or DM_PAPERSIZE
        pDevMode.dmPrintQuality = -1  ' DMRES_DRAFT
        pDevMode.dmTTOption = 3 ' DMTT_SUBDEV
        pDevMode.dmDefaultSource = 0 ' Must be zero
        ' Copy your changes back, then update DEVMODE.
        Call CopyMemory(aDevMode(1), pDevMode, Len(pDevMode))
        
        
        nSize = DocumentProperties(MyhWnd, PrinterHandle, PrinterName, _
                aDevMode(1), aDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)
        
        
        nSize = ResetDC(hPrtDC, aDevMode(1))   ' Reset the DEVMODE for the DC.
        
        ' Close the handle when you are finished with it.
        ClosePrinter (PrinterHandle)
        ' Selection Succeeded! But was Form Added?
End If
        If SelectForm <> FORM_ADDED Then
            SelectForm = FORM_SELECTED
        Else
            SelectForm = FORM_NOT_SELECTED   ' Selection Failed!
        End If
  
  
End Function

Public Sub PrintTest()
    ' Print two test pages to confirm the page size.
    Printer.Print "Top of Page 1."
    Printer.NewPage
    ' Spacing between lines should reflect the chosen page height.
    Printer.Print "Top of Page 2. - Check the page Height (Length.)"
    Printer.EndDoc
    MsgBox "Check Printer " & Printer.DeviceName, vbInformation, "Done!"
End Sub




Public Sub GetCurrentPrinter()
On Error GoTo Err_GetCurrentPrinter
    Dim buffer As String
    Dim i As Integer
        
    'set up the "Size" of the buffer string
    buffer = Space(8192)
    'now get the current default printer the API function GetProfileString returns 1 or 0
    i = GetProfileString("windows", "Device", "", buffer, Len(buffer))
    
    'if we were able to get the profile correctly then continue
    If i Then
        'now parse the returned string
        PreviousPrinterDevName = Mid(buffer, 1, InStr(buffer, ",") - 1)
        PreviousprinterDevDriver = Mid(buffer, InStr(buffer, ",") + 1, InStrRev(buffer, ",") - InStr(buffer, ",") - 1)
        PreviousPrinterDevPort = Mid(buffer, InStrRev(buffer, ",") + 1)
        'set the value of the previous default printer
        PreviousDefaultPrinter = PreviousPrinterDevName
    Else
        PreviousPrinterDevName = ""
        PreviousprinterDevDriver = ""
        PreviousPrinterDevPort = ""
        PreviousDefaultPrinter = ""
    End If
    Exit Sub
Err_GetCurrentPrinter:
    MsgBox ("An error has occured in module GetDefaultPrinter in Function GetCurrentPrinter: The Error Number and Description are " & Err.Number & " " & Err.Description)
    'Resume Err_GetCurrentPrinter_Exit
    
'Err_GetCurrentPrinter_Exit:
 '   Exit Function
    
End Sub

Function PrepareToSetDefaultPrinter(ByVal DeviceName As String) As Boolean
On Error GoTo Err_PrepareToSetDefaultPrinter
    Dim buffer As String
    Dim DriverName As String
    Dim PrinterPort As String
    Dim i As Integer
    Dim iDriver As Integer
    Dim iPort As Integer
    
    If DeviceName <> "" Then
        'get the information for the printer that we just passed from the registry
        buffer = Space(1024)
        i = GetProfileString("PrinterPorts", DeviceName, "", buffer, Len(buffer))
        'now get the DriverName and the PortName out of the Buffer
        'the driver name is the first value in the returned string. Make sure that information was passed back
        iDriver = InStr(buffer, ",")
        If iDriver > 0 Then
            'Ok we know that there was information passed from the registry for the printer that was passed, now get the driver name and port
            DriverName = Left(buffer, iDriver - 1)
            
            'The Port is the second piece of information in the string
            'so we need to make sure that there is a second piece to this string
            iPort = InStr(iDriver + 1, buffer, ",")
            If iPort > 0 Then
                'Now get the Port information
                PrinterPort = Mid(buffer, iDriver + 1, iPort - iDriver - 1)
            End If
        End If
        
        'Now make sure that the DriverName and Port are not blank
        If DriverName <> "" And PrinterPort <> "" Then
            PrepareToSetDefaultPrinter = SetWindows2000DefaultPrinter(DeviceName, DriverName, PrinterPort)
        Else
            PrepareToSetDefaultPrinter = False
        End If
        
    End If
    Exit Function
'Err_PrepareToSetDefaultPrinter
Err_PrepareToSetDefaultPrinter:
    MsgBox ("An error has occured in module GetDefaultPrinter in Function PrepareToSetDefaultPrinter: The Error Number and Description are " & Err.Number & " " & Err.Description)
    'Resume Err_GetCurrentPrinter_Exit
    
Err_PrepareToSetDefaultPrinter_Exit:
    Exit Function
    
End Function

Function SetWindows2000DefaultPrinter(ByVal DeviceName As String, DriverName As String, PrinterPort As String) As Boolean
On Error GoTo Err_SetWindows2000DefaultPrinter
'Ok Set the printer
    Dim DeviceLine As String
    Dim i As Integer
    Dim j As Integer

    DeviceLine = DeviceName & "," & DriverName & "," & PrinterPort

    'In Windows NT and 2000 this information is stored in the registry not in the WIN.INI file
    i = WriteProfileString("windows", "Device", DeviceLine)
    
    If i Then
        j = SendMessage(HWND_BROADCAST, WM_WINICHANGE, 0, "windows")
        SetWindows2000DefaultPrinter = True
        CurrentDefaultPrinter = DeviceName
    Else
        SetWindows2000DefaultPrinter = False
    End If
    Exit Function
Err_SetWindows2000DefaultPrinter:
    MsgBox ("An error has occured in module GetDefaultPrinter in Function SetWindows2000DefaultPrinter: The Error Number and Description are " & Err.Number & " " & Err.Description)
    Resume Err_SetWindows2000DefaultPrinter_Exit
    
Err_SetWindows2000DefaultPrinter_Exit:
    Exit Function

End Function
Public Sub Win95SetDefaultPrinter(PRINTER_NAME As String)
    Dim Handle As Long 'handle to printer
    Dim PrinterName As String
    Dim pd As PRINTER_DEFAULTS
    Dim x As Long
    Dim need As Long ' bytes needed
    Dim pi5 As PRINTER_INFO_5 ' your PRINTER_INFO structure
    Dim LastError As Long

    ' determine which printer was selected
    'PrinterName = List1.List(List1.ListIndex)
    PrinterName = PRINTER_NAME
    ' none - exit
    If PrinterName = "" Then
        Exit Sub
    End If

    ' set the PRINTER_DEFAULTS members
    
    pd.pDatatype = 0&
    pd.DesiredAccess = PRINTER_ALL_ACCESS Or pd.DesiredAccess

    ' Get a handle to the printer
    x = OpenPrinter(PrinterName, Handle, 0&)
    ' failed the open
    If x = False Then
        'error handler code goes here
        Exit Sub
    End If


    x = GetPrinter(Handle, 5, ByVal 0&, 0, need)
    ' don't want to check Err.LastDllError here - it's supposed
    ' to fail  with a 122 - ERROR_INSUFFICIENT_BUFFER
    ' redim t as large as you need
    ReDim T((need \ 4)) As Long

    ' and call GetPrinter for keepers this time
    x = GetPrinter(Handle, 5, T(0), need, need)
    ' failed the GetPrinter
    If x = False Then
        'error handler code goes here
        Exit Sub
    End If

    ' set the members of the pi5 structure for use with SetPrinter.
    ' PtrCtoVbString copies the memory pointed at by the two string
    ' pointers contained in the t() array into a Visual Basic string.
    ' The other three elements are just DWORDS (long integers) and
    ' don't require any conversion
    pi5.pPrinterName = PtrCtoVbString(T(0))
    pi5.pPortName = PtrCtoVbString(T(1))
    pi5.Attributes = T(2)
    pi5.DeviceNotSelectedTimeout = T(3)
    pi5.TransmissionRetryTimeout = T(4)

    ' this is the critical flag that makes it the default printer
    pi5.Attributes = PRINTER_ATTRIBUTE_DEFAULT

    ' call SetPrinter to set it
    x = SetPrinter(Handle, 5, pi5, 0)
    ' failed the SetPrinter
    If x = False Then
        MsgBox "SetPrinterFailed. Error code: " & Err.LastDllError
        Exit Sub
    End If


    ClosePrinter (Handle)

End Sub
Public Function RetrunMyValue()
    GetCurrentPrinter
    ' now that the value has been set... just make sure..
    If PreviousDefaultPrinter <> "" And Len(PreviousDefaultPrinter) > 0 Then
        'ok we have the default printer. now set the default printer to our "special" printer
        'the function PrepareToSetDefaultPrinter is in the GetDefaultPrinter module
        If PrepareToSetDefaultPrinter(SetDefaultPrinterTo) Then
            'assign the properties
            Set prt = Application.Printers(CurrentDefaultPrinter)
            Set Application.Printer = prt
            'MsgBox ("Default Printer has been changed to " & Application.Printer.DeviceName)
            'reopen the report to print
            'DoCmd.OpenReport Reports(ReportToEmail), acViewNormal
            'now that the report has been "printed" reset the current default printer
            'also if the current report was set to use a specific printer rather then the default reset that value as well
            If PrepareToSetDefaultPrinter(PreviousDefaultPrinter) Then
                ' now the current default printer will be the current users previously set default printer
                Set prt = Application.Printers(CurrentDefaultPrinter)
                Set Application.Printer = prt
            End If
            
        End If
          
    End If
    

End Function



Public Sub ParseList(lstCtl As Control, ByVal buffer As String)
    Dim i As Integer
    Dim s As String

    Do
        i = InStr(buffer, Chr(0))
        If i > 0 Then
            s = Left(buffer, i - 1)
            If Len(Trim(s)) Then lstCtl.AddItem s
                buffer = Mid(buffer, i + 1)
            Else
            If Len(Trim(buffer)) Then lstCtl.AddItem buffer
                buffer = ""
            End If

    Loop While i > 0
    
End Sub

Public Function GetAllInstalledPrinters(ByVal buffer As String) As String()

    Dim i As Integer
    Dim iLoop As Integer
    Dim s As String
    Dim ArrName(30) As String
    iLoop = 0
    Do
        i = InStr(buffer, Chr(0))
        If i > 0 Then
            s = Left(buffer, i - 1)
            If Len(Trim(s)) Then ArrName(iLoop) = s
                'lstCtl.AddItem s
                buffer = Mid(buffer, i + 1)
            Else
            If Len(Trim(buffer)) Then ArrName(iLoop) = buffer
                'lstCtl.AddItem Buffer
                buffer = ""
            End If
            iLoop = iLoop + 1
    Loop While i > 0
    
    GetAllInstalledPrinters = ArrName

End Function

Public Sub GetDriverAndPort(ByVal buffer As String, DriverName As String, PrinterPort As String)

    Dim iDriver As Integer
    Dim iPort As Integer

    DriverName = ""
    PrinterPort = ""

    iDriver = InStr(buffer, ",")
    
    If iDriver > 0 Then

        DriverName = Left(buffer, iDriver - 1)
        'The port name is the second entry after the driver name
        'separated by commas.
    
        iPort = InStr(iDriver + 1, buffer, ",")

        If iPort > 0 Then
            'Strip out the port name
            PrinterPort = Mid(buffer, iDriver + 1, _
            iPort - iDriver - 1)
        End If
    
    End If
    
End Sub

Public Function PtrCtoVbString(Add As Long) As String
    Dim sTemp As String * 512, x As Long

    x = lstrcpy(sTemp, Add)
    If (InStr(1, sTemp, Chr(0)) = 0) Then
        PtrCtoVbString = ""
    Else
        PtrCtoVbString = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
    End If
End Function

Public Sub PrintReportInDOSMode(ByVal ReportPath As String, ByVal ReportName As String, ReportTitle As String, ByVal strQueryString As String, ByVal strSelectionFormula As String, PaperSize As String, ShowGroupTree As Boolean, ParamArray Parameters() As Variant)
    'debug.Assert
    Dim iLoop As Integer
    Dim FormIndex As Integer
    Dim SzL As SIZEL
     
    Dim PrnInfo As PrinterInfo
    Dim PSize As String
    'Dim retVal As Integer
    Dim CHK As Boolean
    Dim prhwnd, ret As Long
    Dim SzArr
    Dim PrinterHandle, RetVal As Long
    Dim SizeArr
    Dim Result As Integer
    Dim NumParams As Integer
    Dim MainJob As Integer
    Dim MyPFInfo As PEParameterFieldInfo
    Dim MyVInfo As PEValueInfo
    'Procedurs to Format the ReportPath and ReportName values
    '-----------------------------------------------------------------------------------------------------------
    If Trim(ReportName) = "" Then Exit Sub
    
    If InStr(1, Trim(ReportName), ".", vbTextCompare) <= 0 Then
        ReportName = ReportName & ".rpt"
    End If
    
    If Left$(Trim(ReportName), 1) = "\" Or Left$(Trim(ReportName), 1) = "/" Then
        ReportName = Right$(ReportName, Len(ReportName) - 1)
    End If
    
    If Trim(ReportPath) = "" Then
        ReportPath = App.Path
    End If
    If Right$(Trim(ReportPath), 1) = "\" Or Right$(Trim(ReportPath), 1) = "/" Then
        ReportPath = Left$(ReportPath, Len(ReportPath) - 1)
    End If
    '-----------------------------------------------------------------------------------------------------------
    
    If Trim(PaperSize) = "" Then
        PaperSize = "8.5X12"
    End If
    If Trim(ReportTitle) = "" Then
        reporttile = "Previewing Report : " & ReportName
    Else
    End If
    
    ReportName = ReportPath & "\" & ReportName
    'Get The Generic Printer (if Installed and set it as a default Printer for this Report)
    PrnInfo = GetGenericPrinter(INSTALLED_PRINTERS)
    If Trim(PrnInfo.PrinterName) <> "" Then
            Dim prt As Printer
            For Each prt In Printers
                If prt.DeviceName = PrnInfo.PrinterName Then
                    Set Printer = prt
                    Exit For
                End If
            Next
    Else
            'If Generic Printer is not installed then ask for installation
            Dim ans
            ans = MsgBox("There Is No DOS-Mode Printer Installed on this System" & vbCrLf & _
                   "Do You Want To Install It Now.", vbYesNo, "FACOR ERP : No Text Printer Found")
            If ans = vbYes Then
                Load frmPrintOptions
                frmPrintOptions.Show 1
                Exit Sub
            Else
            End If
    End If
        'Open The Printer and Get its handle
    RetVal = OpenPrinter(Printer.DeviceName, PrinterHandle, 0&)
    'If handel is 0 then there is some problem in the printer
    If RetVal = 0 Then
        MsgBox "Unable to Set DOS-Mode Printer", vbInformation, "FACOR ERP : PRINTER SETTINGS"
    End If
    
    'Set The Papersize (you can also write code to create user defined paper sizes
    'In case you don't know how to do that, please mail me at VB2TheMAX@hotmail.com
    
    PaperSize = Replace(PaperSize, " ", "", 1, Len(PaperSize))
    PaperSize = UCase(Trim(PaperSize))
    
    PSize = GetStringFromList(PaperSize, frmMain.lstPrintForms)
    
    If Trim(PSize) <> "" Then
        SizeArr = Split(PSize, "X", Len(PSize))
    Else
        SizeArr = Split(PaperSize, "X", Len(PaperSize))
    End If
    
    With SzL
        SzL.cx = SizeArr(0) * 2.54 * 10 * 1000
        SzL.cy = SizeArr(1) * 2.54 * 10 * 1000
    End With
    'Get The FormIndex (Each Form has its unique index)
    If Trim(PSize) <> "" Then
        FormIndex = GetFormName(PrinterHandle, SzL, PSize)
    Else
        FormIndex = GetFormName(PrinterHandle, SzL, PaperSize)
    End If
    'If index is 0 then The desired sized form is not availavle, create it
    
    If FormIndex = 0 Then
        MsgBox "Unable to Set the PaperSize : " & PaperSize, vbInformation, "FACOR ERP : PAPER SIZE SETTING"
        'Can write code to create the user defined size form
    Else
    End If
    SetDefaultValues ShowGroupTree, FormIndex
    'set Default values for report (See it very important)
    
    ' Initialize parameter field info and value info structure sizes
    MyPFInfo.StructSize = PE_SIZEOF_PARAMETER_FIELD_INFO
    
    MyVInfo.StructSize = PE_SIZEOF_VALUE_INFO
    
    ' Open print engine
    Result = PEOpenEngine()
    If Result = 0 Then
        MsgBox "Unable to Open Report Engine", vbCritical, "FACOR ERP: VERIFYING REPORT ENGINE"
        Exit Sub
    End If
        
    ' Open the selected report
    MainJob = PEOpenPrintJob(ReportName)
    If MainJob = 0 Then
        MsgBox "An Error Encountered while allocating Print Job" & vbCrLf & _
               "View Report Operation Can not Continued", vbCritical, "FACOR ERP : PRINT JOB ERROR !!"
               Exit Sub
    End If
    ' Get the number of parameters in the report
   ' MsgBox Printer.Port
     iresult = crPESelectPrinter(MainJob, Printer.DriverName, Printer.DeviceName, Printer.Port, Dev)

    If iresult = 0 Then
        MsgBox ShowError(MainJob, "crPESelectPrinter")
    End If
    
'Fetch Nos. of Parameters used in this Report
    NumParams = 0
    NumParams = PEGetNParameterFields(MainJob)
    ' Loop through each parameter field and get and set the current value
    For iLoop = 0 To (NumParams - 1)
    ' Get the selected parameter field value
    Result = PEGetNthParameterField(MainJob, iLoop, MyPFInfo)
    ' Convert the parameter field info structure current value member to a value info structure
        If MyPFInfo.needsCurrentValue = 0 Then 'This param is either linked or not used in Report
        Else
            Result = PEConvertPFInfoToVInfo(MyPFInfo.currentValue, MyPFInfo.valueType, MyVInfo)
            ' Check parameter field value type and place value into appropriate value info structure member
            Select Case MyPFInfo.valueType
        
                Case PE_PF_NUMBER     ' Number parameter
                    MyVInfo.viNumber = Parameters(iLoop)
                Case PE_PF_CURRENCY ' Currency parameter
                    MyVInfo.viCurrency = Parameters(iLoop)
                Case PE_PF_BOOLEAN ' Boolean parameter
                    MyVInfo.viBoolean = Parameters(iLoop)
                Case PE_PF_STRING ' String parameter
                    MyVInfo.viString = Parameters(iLoop)
                Case PE_PF_DATE ' Date parameter
                    MyVInfo.viDate(0) = Year(Parameters(iLoop)) ' 2004
                    MyVInfo.viDate(1) = Month(Parameters(iLoop)) '10
                    MyVInfo.viDate(2) = Day(Parameters(iLoop))
                Case PE_PF_DATETIME ' DateTime Parameter
                    MyVInfo.viDateTime(0) = Year(Parameters(iLoop))
                    MyVInfo.viDateTime(1) = Month(Parameters(iLoop))
                    MyVInfo.viDateTime(2) = Day(Parameters(iLoop))
                    MyVInfo.viDateTime(3) = Hour(Parameters(iLoop))
                    MyVInfo.viDateTime(4) = Minute(Parameters(iLoop))
                    MyVInfo.viDateTime(5) = Second(Parameters(iLoop))
                Case PE_PF_TIME 'Time Parameter
                    MyVInfo.viTime(0) = Hour(Parameters(iLoop))
                    MyVInfo.viTime(1) = Minute(Parameters(iLoop))
                    MyVInfo.viTime(2) = Second(Parameters(iLoop))
            End Select
            ' Convert value info structure to parameter field info current value member
            Result = PEConvertVInfoToPFInfo(MyVInfo, MyPFInfo.valueType, MyPFInfo.currentValue)
            ' Set current value flag
            MyPFInfo.CurrentValueSet = 1
            ' Set the parameter field value
            Result = PESetNthParameterField(MainJob, iLoop, MyPFInfo)
        End If
    Next iLoop
    
    
   
    'Check if Any Selection Formula is supplied
        Dim i As Integer
        Dim txtHand As Long
        Dim Query As String
        Dim RptStrSelectionFormula As String
        Dim RptQueryString As String
        RptStrSelectionFormula = Space(22886)
        RptQueryString = Space(22886)
        'Get the Predefined Selection Formula of Report
        
        i = PEGetSelectionFormula(MainJob, txtHand, Len(RptStrSelectionFormula))
        i = PEGetHandleString(txtHand, RptStrSelectionFormula, Len(RptStrSelectionFormula))
        i = PEGetSQLQuery(MainJob, txtHand, Len(RptQueryString))
        i = PEGetHandleString(txtHand, RptQueryString, Len(RptQueryString))
        ' If User have also supplied the Selection Formula
        If Trim(strSelectionFormula) <> "" Then
            
            If Trim(RptStrSelectionFormula) = "" Then 'If the Report Formula is Blank
                strSelectionFormula = Trim(strSelectionFormula)  'iresult = PESetSelectionFormula(MainJob, strSelectionFormula)
            Else
                strSelectionFormula = Trim(RptStrSelectionFormula) & " AND ( " & Trim(strSelectionFormula) & " )"
                'iresult = PESetSelectionFormula(MainJob, strSelectionFormula)
            End If
        Else 'if No Selection Formula is specified
             strSelectionFormula = Trim(RptStrSelectionFormula)
        End If
    
    If Trim(strQueryString) <> "" Then
        If Trim(RptQueryString) = "" Then
            strQueryString = Trim(strQueryString)
        Else
            strQueryString = Trim(strQueryString)
            
        End If
        
    Else
            strQueryString = Trim(RptQueryString)
    End If
    
    iresult = PESetSQLQuery(MainJob, strQueryString)
    'After Setting the Query String
    'Again Set The Selection Formula for the Report
    'Be aware that query String always overwrites the Selection Formula
    'So set it again
    iresult = PESetSelectionFormula(MainJob, strSelectionFormula)
      
    Handle = PEShowPrintControls(MainJob, 1)
    Handle = PESetWindowOptions(MainJob, WOpt)
    Handle = PESetReportOptions(MainJob, PeReportOpt)
    Handle = PESetAllowPromptDialog(MainJob, False)
    'Do Not Remove Below Commented Lines (For Win 2K users)
    '12045 -- 8625
    Handle = PEOutputToWindow(MainJob, ReportTitle, 0, 0, 0, 0, WS_VISIBLE Or WS_MAXIMIZE Or WS_SYSMENU, 0)
                
    'error handling on the output to window
    If Handle = 0 Then
        MsgBox ShowError(MainJob, "PEOutputToWindow")
    End If
    
    'start the print job
    'CR.Action = 1
    Handle = PEStartPrintJob(MainJob, 1)
    
    If Handle = 0 Then
        MsgBox ShowError(MainJob, "PEOutputToWindow")
    End If
'        RptStrSelectionFormula = Space(22886)
'        RptQueryString = Space(22886)
'        i = PEGetSelectionFormula(MainJob, txtHand, Len(RptStrSelectionFormula))
'        i = PEGetHandleString(txtHand, RptStrSelectionFormula, Len(RptStrSelectionFormula))
'        i = PEGetSQLQuery(MainJob, txtHand, Len(RptQueryString))
'        i = PEGetHandleString(txtHand, RptQueryString, Len(RptQueryString))
'        MsgBox RptStrSelectionFormula
'        MsgBox RptQueryString
    
    ReportPath = ""
    ReportName = ""
    ReportTitle = ""
    strQueryString = ""
    strSelectionFormula = ""
    PaperSize = ""
    RptStrSelectionFormula = ""
    RptQueryString = ""
       
    
    ' Close the print job
    PEClosePrintJob (MainJob)
    
    ' Close the engine
    'This will Also close the Report
    'PECloseEngine
    
End Sub
Public Function ShowError(JobNum As Integer, Optional ModuleName As String = "", Optional CloseJob As Boolean = True) As String
    Dim TextHand, RetVal As Long
    Dim TextSize As String * 255
    Dim buffer As String * 255
    Dim ReturnString As String * 500
    
    ErrorNum = PEGetErrorCode(JobNum)
    RetVal = PEGetErrorText(JobNum, TextHand, 1084)
    RetVal = PEGetHandleString(TextHand, TextSize, 255)
    
    If Trim(ModuleName) <> "" Then
        ReturnString = "An Error Occurd in Module [ " & ModuleName & " ] With No. " & ErrorNum & vbCrLf & _
                   "The Error Says That : " & vbCrLf & vbCrLf & TextSize
    Else
        ReturnString = "An Error Occurd With No. " & ErrorNum & vbCrLf & _
                   "The Error Says That : " & vbCrLf & vbCrLf & TextSize
    End If
    
    If CloseJob = True Then
        RetVal = PEClosePrintJob(JobNum)
    End If
    
    ShowError = ReturnString
    
           
End Function
'Public Function SetReportPrintingFormat(ParentForm As Form, _
'                CRView As CrystalReport, _
'                strFullReportName As String, _
'                Optional strFullReportPath As String = "\\main\e\Facorexe\Report", _
'                Optional ConnectionString As String = "", _
'                Optional LogOnServerInfo As String = "", _
'                Optional ShowGroupTree As Boolean = False, _
'                Optional PaperSize As String = "14X12") _
'                As Boolean
'    On Error GoTo Errhandler
'    Dim PrnInfo As PrinterInfo
''        ConnectionString = "driver={SQL Server};server=192.0.0.03;" & _
''      "database=new_facor;uid=sa;pwd="
''
''        .Connect = cnstr
''
''        .LogOnServer "p2ssql.dll", "192.0.0.03", _
''                     "new_facor", "sa", ""
'
''-------------------------------------------------------------------------
''Validate Input Parameters
''-------------------------------------------------------------------------
'If Right$(strFullReportName, 1) = "\" Or Right$(strFullReportName, 1) = "/" Then
'    strFullReportName = Left$(strFullReportName, Len(strFullReportName) - 1)
'End If
'
'If Right$(strFullReportPath, 1) = "\" Or Right$(strFullReportPath, 1) = "/" Then
'    strFullReportPath = Left$(strFullReportPath, Len(strFullReportPath) - 1)
'End If
'
''-------------------------------------------------------------------------
''-------------------------------------------------------------------------
'
'    With CRView
'
'    If Trim(ConnectionString) <> "" Then
'       .Connect = ConnectionString
'    End If
'    If Trim(LogOnServerInfo) <> "" Then
'       ' .LogOnServer LogOnServerInfo
'
'    End If
'
''-------------------------------------------------------------------------
''Reset Previous Values, if any and Prepare ReportViewer For New Entry
''-------------------------------------------------------------------------
'        .Reset
'        .SQLQuery = ""
'        .ReportFileName = ""
''-------------------------------------------------------------------------
''-------------------------------------------------------------------------
'        .ReportFileName = strFullReportPath & "\" & strFullReportName
'        .WindowState = crptMaximized
'        .WindowShowSearchBtn = True
'        .WindowAllowDrillDown = True
'        .WindowShowRefreshBtn = True
'        .WindowShowPrintSetupBtn = True
'        .WindowShowPrintBtn = True
'        .WindowShowGroupTree = ShowGroupTree
'        .WindowShowExportBtn = True
'        .WindowState = crptMaximized
'        .WindowBorderStyle = crptFixedSingle
'        .WindowTitle = strWindowTitle
'        .Destination = crptToWindow
'
'        PrnInfo = GetGenericPrinter(INSTALLED_PRINTERS)
'
'        If Trim(PrnInfo.PrinterName) <> "" Then
'            .PrinterDriver = PrnInfo.PrinterDriver
'            .PrinterName = PrnInfo.PrinterName
'            .PrinterPort = PrnInfo.PrinterPort
'            Dim prt As Printer
'            For Each prt In Printers
'                If prt.DeviceName = .PrinterName Then
'                    Set Printer = prt
'                    Exit For
'                End If
'            Next
'        Else
'            Dim ans
'            ans = MsgBox("There Is No DOS-Mode Printer Installed on this System" & vbCrLf & _
'                   "Do You Want To Install It Now.", vbOKOnly, "FACOR ERP : No Text Printer Found")
'            If ans = vbYes Then
'                Load frmPrintOptions
'                frmPrintOptions.Show 1
'            Else
'            End If
'            .PrinterDriver = PreviousprinterDevDriver
'            .PrinterName = PreviousPrinterDevName
'            .PrinterPort = PreviousPrinterDevPort
'        End If
'        Dim PSize As String
'        Dim RetVal As Integer
'        Dim CHK As Boolean
'        Dim prhwnd, ret As Long
'        Dim SzArr
'
'        ret = OpenPrinter(Printer.DeviceName, prhwnd, 0&)
'        Dim SzL As SIZEL
'
'            PSize = GetStringFromList(PaperSize, MDIStore.lstPrintForms)
'
'
'       End With
'
''            SzArr = Split(UCase(PaperSize), "X", Len(PaperSize))
''            'RetVal = SetPrinterAttributes(ParentForm.hwnd, SzArr(0), SzArr(1), 0)
''            SzL.cx = SzArr(0) * 2.54 * 10 * 1000
''            SzL.cy = SzArr(1) * 2.54 * 10 * 1000
'            RetVal = SelectForm(PSize, ParentForm.HWnd)
''
''            chk = SetAsDefaultForm(prhwnd, SzL, PSize)
'''            ShowPrinterSetup ParentForm, SzL
''
''            'B = SetWindows2000DefaultPrinter(Printer.DeviceName, Printer.DriverName, "LPT1:")
''
''            'MsgBox DM.dmDeviceName
''            'retVal = SetPrinterAttributes(ParentForm.hWnd, SzArr(0), SzArr(1), 0)
''            Select Case RetVal
''
''                Case FORM_NOT_SELECTED   ' 0
''                     ' Selection failed!
''                        'MsgBox "Unable to retrieve PaperSize" & vbCrLf & _
''                                "Please Select Paper Size : " & PaperSize & " Manually", vbExclamation, "FACOR ERP : FORM SIZE SELECTION FAILED!"
''                        SetReportPrintingFormat = False
''
''                Case FORM_SELECTED   ' 1
''                     ' Selection succeeded!
''                        'PrintTest     ' Comment this line to avoid printing
''                      SetReportPrintingFormat = True
''                Case FORM_ADDED   ' 2
''                    ' Form added and selected.
''                    'List1.Clear   ' Reflect the addition in the ListBox
''                    'Form_Load     ' by rebuilding the list.
''                    SetReportPrintingFormat = True
''                End Select
''      End With
'        'SetReportPrintingFormat = True
'    Exit Function
'Errhandler:
'
'    MsgBox "An Error Encounterd While Connecting to Report Server" & vbCrLf & _
'           "Module Name : SetReportPrintingFormat" & vbCrLf & _
'           "Error Description: " & Err.Description, vbCritical, "FACOR ERP"
'    SetReportPrintingFormat = False
'    Exit Function
'End Function

Public Function SetPrinterAttributes(MyhWnd As Long, ByVal FormWidth As Double, ByVal FormHeight As Double, Optional PaperCut As Integer = 0) As Integer

    Dim nSize As Long           ' Size of DEVMODE
    
    Dim pDevMode As PSDEVMODE
    
    Dim PrinterHandle As Long   ' Handle to printer
    
    Dim hPrtDC As Long          ' Handle to Printer DC
    
    Dim PrinterName As String
    
    Dim aDevMode() As Byte      ' Working DEVMODE
    
     
    
    PrinterName = Printer.DeviceName  ' Current printer
    
    hPrtDC = Printer.hdc              ' hDC for current Printer
    
     
    
    ' Get a handle to the printer.
    
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
    
        ' Retrieve the size of the DEVMODE.
    
        nSize = DocumentProperties(MyhWnd, PrinterHandle, PrinterName, 0&, _
                0&, 0&)
    
        ' Reserve memory for the actual size of the DEVMODE.
    
        ReDim aDevMode(1 To nSize)
    
     
    
        ' Fill the DEVMODE from the printer.
    
        nSize = DocumentProperties(MyhWnd, PrinterHandle, PrinterName, _
                aDevMode(1), 0&, DM_OUT_BUFFER)
    
        ' Copy the Public (predefined) portion of the DEVMODE.
    
        Call CopyMemory(pDevMode, aDevMode(1), Len(pDevMode))
    
        
    
        ' Change the appropriate member in the DevMode.
    
        pDevMode.dmPublic.dmFields = pDevMode.dmPublic.dmFields Or _
            DM_PAPERSIZE Or DM_PAPERLENGTH Or DM_PAPERWIDTH
    
        pDevMode.dmPublic.dmFields = pDevMode.dmPublic.dmFields And (Not DM_FORMNAME)
    
        ' this needs to be 256 + (formlength) + (0 or 1) where 0=die cut 1=continuous
    
        ' the example here sets form length (height) to 1" = 25.4mm
        pDevMode.dmPublic.dmPaperSize = 256 + (254) + 1
    
        ' this needs to be the form length again (in tenths of mm)
    
        pDevMode.dmPublic.dmPaperLength = (FormHeight * 2.54) * 10   '101.8
    
        ' the form width in tenths of mm
    
        pDevMode.dmPublic.dmPaperWidth = (FormWidth * 2.54) * 10
    
        
    
        ' Copy your changes back, then update DEVMODE.
    
        Call CopyMemory(aDevMode(1), pDevMode, Len(pDevMode))
    
        nSize = DocumentProperties(MyhWnd, PrinterHandle, PrinterName, _
                aDevMode(1), aDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)
    
     
    
        nSize = ResetDC(hPrtDC, aDevMode(1))   ' Reset the DEVMODE for the DC.
    
     
    
        ' Close the handle when you are finished with it.
    
        ClosePrinter (PrinterHandle)
    
    End If
    
End Function




Public Sub SetPrinterSettings()
    Dim buffer As String
    Dim osinfo As OSVERSIONINFO
    Dim retvalue As Integer
    Dim iLoop As Integer
    
    buffer = Space(8192)
    r = GetProfileString("PrinterPorts", vbNullString, "", _
    buffer, Len(buffer))
    INSTALLED_PRINTERS = GetAllInstalledPrinters(buffer)
    
    
    
    GetCurrentPrinter
End Sub
    
Public Function IsGENERICPritnerInstalled(PrnArr() As String) As Boolean
    Dim iLoop As Integer
    Dim iLen As Integer
    iLen = Len("GENERIC")
    
    For iLoop = 0 To UBound(PrnArr)
        If InStr(1, UCase(PrnArr(iLoop)), UCase("GENERIC")) > 0 Then
           IsGENERICPritnerInstalled = True
           Exit Function
        Else
            IsGENERICPritnerInstalled = False
        End If
        
    Next iLoop
End Function

Public Function GetGenericPrinter(PrnArr As Variant) As PrinterInfo
    Dim iLoop As Integer
    Dim lVal As Long
    Dim chkPrinter As Boolean
    Dim TextPrinter, buffer As String
    Dim DriverName As String
    Dim PortName As String
    DriverName = ""
    PortName = ""
    buffer = ""
    buffer = Space(8192)
    Dim PrnInfo As PrinterInfo
    With PrnInfo
        .PrinterDevice = ""
        .PrinterDriver = ""
        .PrinterName = ""
        .PrinterPort = ""
    End With
    TextPrinter = ""
    chkPrinter = False
    For iLoop = 0 To UBound(PrnArr)
        If InStr(1, UCase(PrnArr(iLoop)), UCase("GENERIC")) > 0 And _
           InStr(1, UCase(PrnArr(iLoop)), UCase("TEXT")) > 0 Then
           chkPrinter = True
           TextPrinter = PrnArr(iLoop)
           Exit For
        Else
            chkPrinter = False
        End If
        
    Next iLoop
    
    
    If chkPrinter = False Then
        GetGenericPrinter = PrnInfo
    Else
        lVal = GetProfileString("PrinterPorts", TextPrinter, "", _
                             buffer, Len(buffer))
        GetDriverAndPort buffer, DriverName, PortName
        With PrnInfo
            .PrinterDriver = DriverName
            .PrinterPort = PortName
            .PrinterName = TextPrinter
        End With
        GetGenericPrinter = PrnInfo
    End If
    
End Function


'*********************************************************************************************
' Adds a new printer to the system
'
' Parameters:
'
'   PrinterName: The friendly name of the new printer
'   Driver:      The driver name
'   Port:        Optional. The port where the printer is connected. Default is "LPT1:"
'   Server:      Optional. The name of printer server. Default is local printer.
'   Comment:     Optional. Any comment you want to add to the printer
'   PhysicalLocation: Optional. Where the printer is located.
'
'*********************************************************************************************
Public Sub AddNewPrinter(ByVal PrinterName As String, _
                         ByVal Driver As String, _
                         Optional ByVal Port As String = "LPT1:", _
                         Optional Server As String, _
                         Optional Comment As String, _
                         Optional PhysicalLocation As String, Optional ShareName As String = "GenericT")
Dim hPrint As Long, PI As PRINTER_INFO_2

    ' Fill the PRINTER_INFO_2 struct
With PI
        .pServerName = Server
        .pPrinterName = PrinterName
        .pDriverName = Driver
        .pPortName = Port
        .pPrintProcessor = "WinPrint"
        .Priority = 1
        .DefaultPriority = 1
        .pComment = Comment
        .pDatatype = "RAW"
        .pShareName = ShareName
    End With
    
    ' Add the printer
    hPrint = AddPrinter(vbNullString, 2, PI)
    ' Raise an error if the printer
    ' cannot be created
    If hPrint = 0 Then
         Dim ErrNum As Long
         ErrNum = Err.LastDllError
        If ErrNum = 1797 Then 'Driver is Unknown
           AddNewPrinter PrinterName, "winspool"
        ElseIf ErrNum = 1798 Then 'Print Processor is Unknown
            RaiseAPIError "AddNewPrinter"
        ElseIf ErrNum = 1802 Then 'Printer Already Exists
            MsgBox "Printer Driver Updated, Please Re-View the Report", vbInformation, "FACOR ERP: PRINTER UPGRADATION "
        End If
    Else
            MsgBox "DOS-Mode Printer Installed Successfully, Please Re-View the Report", vbInformation, "FACOR ERP: PRINTER UPGRADATION "
            Call frmMain.LoadDetails
            
    End If
    ' Close the printer handle
    ClosePrinter hPrint
    
End Sub

'*********************************************************************************************
' Adds a new printer driver
'
' Parameters:
'
'   DriverName:  Specifies the name of the driver.
'   DriverFile:  Specifies a filename or full path and filename for the file
'                that contains the device driver
'   Server:      Optional. The name of printer server. Default is local
'   ConfigFile:  Optional. Full path to the config file. Default is no config file
'   PhysicalLocation: Optional. Where the printer is located
'
'*********************************************************************************************
Public Sub AddNewDriver(ByVal DriverName As String, ByVal DriverFile As String, ByVal DataFile As String, Optional ByVal ConfigFile As String, Optional Server As String = vbNullString)
Dim hPrint As Long, DI As DRIVER_INFO_2

     ' If not ConfigFile is specified
     ' use DriverFile
     If ConfigFile = "" Then ConfigFile = DriverFile

     ' Fill the DRIVER_INFO_2 struct
     With DI
         .pDriverPath = DriverFile
         .pName = DriverName
         .pDataFile = DataFile
         .pConfigFile = DriverFile    ' Changed DriverFile to ConfigFile
     End With

     ' Add the printer driver
     'hPrint = AddPrinterDriver(Server, 2, DI)

     ' Raise an error if the driver
     ' cannot be added
    ' If hPrint = 0 Then RaiseAPIError "AddNewDriver"

End Sub

Public Function RaiseAPIError(ByVal Source As String)
Dim ErrorMsg As String, ErrNum As Long

    ErrNum = Err.LastDllError

    ErrorMsg = String(256, 0)
    ErrorMsg = Left$(ErrorMsg, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, ErrNum, 0&, ErrorMsg, Len(ErrorMsg), ByVal 0))

    MsgBox "The Error : " & ErrNum & "  Raised By : " & Source & vbCrLf & "States That " & vbCrLf & ErrorMsg

End Function

Public Function GetPrintDriverPath() As String
    Dim hpHandle As Long
    Dim lError As Long
    Dim lResult As Long
    Dim sDriverBuf As String * MAX_PATH
    Dim sDriverPath As String
    Dim lPathSize As Long
    
    
    ' First we get the folder for the printer driver files
    lResult = GetPrinterDriverDirectory(vbNullString, vbNullString, 1, _
                                        sDriverBuf, MAX_PATH, lPathSize)
    If lResult = 0 Then
      'ErrorHandler "Unable to retrieve the printer driver folder.",
    'Err.LastDllError
    Else
      lResult = InStr(sDriverBuf, vbNullChar)
      sDriverPath = Left$(sDriverBuf, lResult - 1)
      If (Right$(sDriverPath, 1) <> "\") Then sDriverPath = sDriverPath & "\"
      GetPrintDriverPath = sDriverPath
    End If

End Function

Public Sub ShowPrinterSetup(frmOwner As Form, PaperSize As SIZEL)
       Dim PRINTSETUPDLG As PRINTSETUPDLG_TYPE
       Dim DevMode As DEVMODE_TYPE
       Dim DevName As DEVNAMES_TYPE
       Dim PI As POINTAPI
             
       Dim lpDevMode As Long, lpDevName As Long
       Dim bReturn As Integer
       Dim objPrinter As Printer, NewPrinterName As String
       Dim strSetting As String
   
       ' Use PrintDialog to get the handle to a memory
       ' block with a DevMode and DevName structures

       PRINTSETUPDLG.lStructSize = Len(PRINTSETUPDLG)
       PRINTSETUPDLG.hwndOwner = frmOwner.hwnd
       PRINTSETUPDLG.ptPaperSize.x = PaperSize.cx
       PRINTSETUPDLG.ptPaperSize.y = PaperSize.cy
       ' Set the current orientation and duplex setting
       DevMode.dmDeviceName = Printer.DeviceName
       DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX _
          Or DM_COPIES Or DM_PAPERSIZE
       DevMode.dmOrientation = Printer.Orientation
       DevMode.dmCopies = Printer.Copies
       DevMode.dmPaperSize = 138
       'DevMode.dmPaperLength = PaperSize.cy
       'DevMode.dmPaperWidth = PaperSize.cx
    
       On Error Resume Next
       DevMode.dmDuplex = Printer.Duplex
        
       On Error GoTo 0
   
       ' Allocate memory for the initialization hDevMode structure
       ' and copy the settings gathered above into this memory
       PRINTSETUPDLG.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or _
          GMEM_ZEROINIT, Len(DevMode))
       lpDevMode = GlobalLock(PRINTSETUPDLG.hDevMode)
       If lpDevMode > 0 Then
          CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
           bReturn = GlobalUnlock(PRINTSETUPDLG.hDevMode)
       End If
   
       ' Set the current driver, device, and port name strings
       With DevName
           .wDriverOffset = 8
           .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
           .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
           .wDefault = 0
       End With
       With Printer
           DevName.extra = .DriverName & Chr(0) & _
           .DeviceName & Chr(0) & .Port & Chr(0)
       End With
   
       ' Allocate memory for the initial hDevName structure
       ' and copy the settings gathered above into this memory
       PRINTSETUPDLG.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or _
           GMEM_ZEROINIT, Len(DevName))
       lpDevName = GlobalLock(PRINTSETUPDLG.hDevNames)
       If lpDevName > 0 Then
           CopyMemory ByVal lpDevName, DevName, Len(DevName)
           bReturn = GlobalUnlock(lpDevName)
       End If
   
       ' Call the print dialog up and let the user make changes
       'MsgBox DevMode.dmPaperSize
       
       If PageSetupDialog(PRINTSETUPDLG) Then
   
           ' First get the DevName structure.
           lpDevName = GlobalLock(PRINTSETUPDLG.hDevNames)
               CopyMemory DevName, ByVal lpDevName, 45
           bReturn = GlobalUnlock(lpDevName)
           GlobalFree PRINTSETUPDLG.hDevNames
   
           ' Next get the DevMode structure and set the printer
           ' properties appropriately
           lpDevMode = GlobalLock(PRINTSETUPDLG.hDevMode)
               CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
           bReturn = GlobalUnlock(PRINTSETUPDLG.hDevMode)
           GlobalFree PRINTSETUPDLG.hDevMode
           NewPrinterName = UCase$(Left(DevMode.dmDeviceName, _
               InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
          ' NewPrinterName = Printer.DeviceName
           If UCase(Printer.DeviceName) <> UCase(NewPrinterName) Then
               For Each objPrinter In Printers
                  If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                       Set Printer = objPrinter
                       
                  End If
               Next
           End If
           On Error Resume Next
   
           ' Set printer object properties according to selections made
          ' by user
           DoEvents
           With Printer
                
               .Copies = DevMode.dmCopies
               .Duplex = DevMode.dmDuplex
               .Orientation = DevMode.dmOrientation
           End With
           On Error GoTo 0
       End If
   
       ' Display the results in the immediate (debug) window
       With Printer
           If .Orientation = 1 Then
               strSetting = "Portrait.  "
           Else
               strSetting = "Landscape. "
           End If
           Debug.Print "Copies = " & .Copies, "Orientation = " & _
              strSetting & GetDuplex(Printer.Duplex)
       End With
   End Sub
Function GetDuplex(lDuplex As Long) As String
        Dim TempStr As String
              
        If lDuplex = DMDUP_SIMPLEX Then
           TempStr = "Duplex is turned off (1)"
        ElseIf lDuplex = DMDUP_VERTICAL Then
           TempStr = "Duplex is set to VERTICAL (2)"
        ElseIf lDuplex = DMDUP_HORIZONTAL Then
           TempStr = "Duplex is set to HORIZONTAL (3)"
        Else
           TempStr = "Duplex is set to undefined value of " & lDuplex
        End If
        GetDuplex = TempStr   ' Return descriptive text
 End Function

Private Function SetAsDefaultForm(ByVal PrinterHandle As Long, FormSize As SIZEL, FormName As String) As Boolean
Dim FI1 As sFORM_INFO_1
Dim aFI1() As Byte
Dim RetVal As Long

With FI1
    .Flags = 0
    .pName = FormName
    With .Size
        .cx = FormSize.cx
        .cy = FormSize.cy
    End With
    With .ImageableArea
        .Left = 0
        .Top = 0
        .Right = FI1.Size.cx
        .Bottom = FI1.Size.cy
    End With
End With
ReDim aFI1(Len(FI1))
Call CopyMemory(aFI1(0), FI1, Len(FI1))
RetVal = SetForm(PrinterHandle, FormName, 1, aFI1(0))
If RetVal = 0 Then
    If Err.LastDllError = 5 Then
        MsgBox "You do not have permissions to add a form to " & _
           Printer.DeviceName, vbExclamation, "Access Denied!"
    Else
        MsgBox "Error: " & Err.LastDllError, "Error Adding Form"
    End If
    SetAsDefaultForm = False
Else
    SetAsDefaultForm = True
End If
End Function

  
Public Function ValidateForms(FormName As String) As Boolean
    Dim strTemp As String
    FormName = Trim(FormName)
    strTemp = Replace(FormName, " ", "", 1, Len(FormName))
    strTemp = UCase(strTemp)
    With PrinterFormNames
        If strTemp = .f10X12 Or _
           strTemp = .f10X6 Or _
           strTemp = .f14X12 Or _
           strTemp = .f14X6 Or _
           strTemp = .f14X14 Or _
           strTemp = .f14X15 Or _
           strTemp = .f8_5X12 Then
           ValidateForms = True
        Else
            ValidateForms = False
        End If
    End With
End Function
Public Function GetStringFromList(LIKEClause As String, lstBox As ListBox) As String
    Dim iLen As Integer
    Dim iLoop As Integer
    Dim FullVal As String
    iLen = Len(LIKEClause)
    For iLoop = 0 To lstBox.ListCount - 1
        If InStr(1, UCase(lstBox.List(iLoop)), UCase(LIKEClause)) > 0 Then
            FullVal = lstBox.List(iLoop)
            GetStringFromList = FullVal
            Exit Function
        End If
    
    Next iLoop
    GetStringFromList = ""

End Function


Public Sub SetStringInList(strToSet As String, lstBox As ListBox)
    Dim iLen As Integer
    Dim iLoop As Integer
    Dim FullVal As String
    iLen = Len(strToSet)
    For iLoop = 0 To lstBox.ListCount - 1
        If InStr(1, UCase(lstBox.List(iLoop)), UCase(strToSet)) > 0 Then
            'FullVal = lstBox.List(iLoop)
            'GetStringFromList = FullVal
            lstBox.ListIndex = iLoop
            Exit Sub
        End If
    
    Next iLoop
    
    If lstBox.ListCount < 0 Then
    Else
        lstBox.ListIndex = 0
    End If

End Sub




'Public Sub UseForm(FormName As String)
'Dim RetVal As Integer
'
'RetVal = SelectForm(FormName, Me.HWnd)
'Select Case RetVal
'    Case FORM_NOT_SELECTED   ' 0
'        ' Selection failed!
'        MsgBox "Unable to retrieve From name", vbExclamation, _
'           "Operation halted!"
'    Case FORM_SELECTED   ' 1
'        ' Selection succeeded!
'        PrintTest     ' Comment this line to avoid printing
'    Case FORM_ADDED   ' 2
'        ' Form added and selected.
'        List1.Clear   ' Reflect the addition in the ListBox
'        GetPrinterFormNames     ' by rebuilding the list.
'End Select
'End Sub

