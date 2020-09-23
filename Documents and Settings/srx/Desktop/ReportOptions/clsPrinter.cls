VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************************************
'
' .INF file parser class
'
'*********************************************************************************************
'          The class check for NT to
'          lauch the .NT section from
'          the .inf file
'*********************************************************************************************
Option Explicit

Dim hInf As Long

Enum InfErrors
    ieFileNotOpen = vbObjectError
    ieCannotOpen
End Enum
'Enum InfStyles
'    INF_STYLE_NONE = 0       ' unrecognized or non-existent
'    INF_STYLE_OLDNT = 1      ' winnt 3.x
'    INF_STYLE_WIN4 = 2       ' Win95
'End Enum

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long
'*********************************************************************************************
' Appends a new .inf file to the current
' loaded .inf
'
' Parameters:
'
'   FileName:   File name of .inf file
'*********************************************************************************************
Public Sub AppendInf(ByVal FileName As String)
Dim ErLine As Long

    If hInf > 0 Then
        SetupOpenAppendInfFile FileName, hInf, ErLine
    Else
        Err.Raise ieFileNotOpen, , "Open a .inf file before append"
    End If

End Sub

'*********************************************************************************************
' Closes the .inf file
'*********************************************************************************************
Public Sub CloseInf()

    SetupCloseInfFile hInf
    hInf = 0

End Sub


'*********************************************************************************************
' Enumerates keys under a section.
'
' Parameters:
'
'   Section:    Parent section.
'   EnumValues: Target array of keys
'*********************************************************************************************
Public Sub EnumSectionKeys(ByVal Section As String, EnumValues() As String)
Dim Lines As Long, CTX As INFCONTEXT, Req As Long

    ' Clear the array
    Erase EnumValues

    ' Get section line count
    Lines = SetupGetLineCount(hInf, Section)

    ' Check if there are lines
    ' to read
    If Lines > 0 Then

        ' Redim the array
        ReDim EnumValues(0 To Lines - 1)

        ' Enumerate fields
        For Lines = 0 To Lines - 1

            ' Get the line
            SetupGetLineByIndex hInf, Section, Lines, CTX

            ' Get the field length
            SetupGetStringField CTX, 0, vbNullString, 0, Req

            ' Initialize buffer
            EnumValues(Lines) = String$(Req, 0)

            ' Get the first field (0). The first field is
            ' the key name
            SetupGetStringField CTX, 0, EnumValues(Lines), Req, Req

        Next

    End If

End Sub

'*********************************************************************************************
' Returns a Key value
'
' Parameters:
'
'   Section:    Parent section
'   Key:        Key name
'*********************************************************************************************
Public Function GetValue(ByVal Section As String, ByVal Key As String) As String
Dim Req As Long

    ' Get the field length
    SetupGetLineText ByVal 0&, hInf, Section, Key, vbNullString, 0, Req

    If Req > 0 Then

        ' Initialize buffer
        GetValue = String$(Req, 0)

        ' Get the text
        SetupGetLineText ByVal 0&, hInf, Section, Key, GetValue, Req, Req

        ' Remove the trailing null
        GetValue = Left$(GetValue, Len(GetValue) - 1)

    End If

End Function

'*********************************************************************************************
' Returns a key field.
'
' Parameters:
'
'   Section:    Parent section
'   Key:        Key name
'   Field:      Field index. Field are separated by ",".
'*********************************************************************************************
Public Function GetField(ByVal Section As String, ByVal Key As String, ByVal Field As Long) As String
Dim Req As Long, CTX As INFCONTEXT
    
    ' Find the line
    SetupFindFirstLine hInf, Section, Key, CTX

    ' Get the field length
    SetupGetStringField CTX, Field, vbNullString, 0, Req
    'MsgBox CTX.CurrentInf
    If Req > 0 Then

        ' Initialize buffer
        GetField = String$(Req, 0)

        ' Get the text
        SetupGetStringField CTX, Field, GetField, Req, Req

        ' Remove trailing null
        GetField = Left$(GetField, Len(GetField) - 1)

    End If

End Function

'*********************************************************************************************
' Executes an install section
'
' Parameters:
'
'   Section:    Section to execute
'   Flags:      Optional. Specifies what keys the installer
'               will execute. Default is all keys.
'   CopyFlags:  Optional. Specifies the copy flags. Default is SP_COPY_NEWER
'   OwnerhWnd:  Specifies the parent window. Default is the desktop.
'*********************************************************************************************
Public Function InstallSection(ByVal Section As String, Optional ByVal Flags As InstallFlags = &H1F, Optional ByVal CopyFlags As CopyFlags = &H4, Optional ByVal OwnerhWnd As Long) As Boolean
Dim hMod As Long, DefaultCTX As Long
Dim DefaultCallback As Long

    ' Open setupapi to get
    ' SetupDefaultQueueCallback address
    hMod = LoadLibrary("setupapi.dll")

    ' Get SetupDefaultQueueCallback address
    DefaultCallback = GetProcAddress(hMod, "SetupDefaultQueueCallback")

    ' Get the default context
    DefaultCTX = SetupInitDefaultQueueCallback(OwnerhWnd)

    ' Check for NT
    If (GetVersion And &H80000000) = 0 Then

        ' Check if section already
        ' includes ".NT" at the end
        If UCase$(Right$(Section, 3)) <> ".NT" Then
            Section = Section & ".NT"
        End If

    End If

    ' Install the section
    InstallSection = SetupInstallFromInfSection(OwnerhWnd, hInf, Section, Flags, 0&, vbNullString, CopyFlags, DefaultCallback, DefaultCTX, 0&, ByVal 0&)

    ' Unload setupapi.dll
    FreeLibrary hMod

End Function





'*********************************************************************************************
' Opens a .inf file
'
' Parameters:
'
'   FileName:   .inf file name
'   Class:      Optional. The class to install. Default is BASE (software).
'*********************************************************************************************
Public Sub OpenInf(ByVal FileName As String, Optional ByVal Class As String = "Base")
Dim ErLine As Long

    ' Close the previous inf file
    If hInf Then CloseInf

    ' Open the inf file
    '2 = WIN 9X
    hInf = SetupOpenInfFile(FileName, Class, 2, ErLine)

    ' The inf file cannot be opened
    ' maybe the file name or the
    ' class is wrong
    If hInf = -1 Then Err.Raise ieCannotOpen, , "The .inf file cannot be opened. Check the file name and .inf class."

End Sub


Private Sub Class_Terminate()

    SetupCloseInfFile hInf

End Sub



