VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMain 
      Caption         =   "View Report"
      Height          =   720
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   4605
      Begin VB.CommandButton Command1 
         Caption         =   "View Report in Crystal Format"
         Height          =   450
         Left            =   1185
         TabIndex        =   7
         Top             =   195
         Width           =   2595
      End
   End
   Begin VB.PictureBox picContainer 
      Height          =   225
      Left            =   5310
      ScaleHeight     =   165
      ScaleWidth      =   6510
      TabIndex        =   2
      Top             =   5190
      Width           =   6570
      Begin VB.ListBox lstPrintForms 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   -15
         TabIndex        =   5
         Top             =   15
         Width           =   1800
      End
      Begin VB.ListBox lstPrinters 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   1905
         TabIndex        =   4
         Top             =   15
         Width           =   2010
      End
      Begin VB.TextBox txtWindowsLogger 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9480
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   270
         Width           =   1560
      End
   End
   Begin VB.ComboBox cboAge 
      Height          =   315
      Left            =   3015
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   30
      Width           =   1590
   End
   Begin VB.Label lblAge 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select Age From the Combo"
      Height          =   285
      Left            =   30
      TabIndex        =   1
      Top             =   75
      Width           =   2880
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CMD As New ADODB.Command
Dim PM As New ADODB.Parameter
Dim DummyRS As New ADODB.Recordset

Private Sub Command1_Click()
    Screen.MousePointer = 11
        PrintReportInDOSMode App.Path, "rptMyReport.rpt", "Displaying Records for Age " & cboAge.Text, "", "", "8.5X12", False, CInt(cboAge.Text)
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()
    DoEvents
      btnFlatAll Me
      Me.Caption = "Crystal Report View / Print in Draft Style"
    DoEvents
End Sub

Private Sub Form_Load()
    With CMD
            .ActiveConnection = CON
            .CommandType = adCmdText
            .CommandText = "SELECT distinct tblEmployee.Age FROM tblEmployee;"
            Set DummyRS = .Execute
    End With
    
    PopulateComboBoxWithDefinedIndex cboAge, DummyRS, , 0, True
    DoEvents
    
    LoadDetails
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set CON = Nothing
    CMD.Parameters.Refresh
    Set CMD = Nothing
    
    Set PM = Nothing
    Set DummyRS = Nothing
End Sub


Public Sub LoadDetails()

        With PrinterFormNames
                .f10X12 = "10X12"
                .f10X6 = "10X6"
                .f14X12 = "14X12"
                .f14X6 = "14X6"
                .f8_5X12 = "8.5X12"
                .f14X14 = "14X14"
                .f14X15 = "14X15"
            End With
            Dim PRTNAME As String
             SetPrinterSettings ' Must be called before For Loop
            For iLoop = 0 To UBound(INSTALLED_PRINTERS)
                If Trim(INSTALLED_PRINTERS(iLoop)) = "" Then
                Else
                    PRTNAME = Trim(INSTALLED_PRINTERS(iLoop))
                    PRTNAME = Replace(PRTNAME, " ", "", 1, Len(PRTNAME))
                    PRTNAME = UCase(PRTNAME)
                    lstPrinters.AddItem PRTNAME, iLoop
                End If
            Next iLoop
           'Set WSHShell = CreateObject("WScript.Shell")
            GetPrinterFormNames
          

End Sub
Public Sub GetPrinterFormNames()
    Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1           ' Working FI1 array
    Dim Temp() As Byte                  ' Temp FI1 array
    
    
    Dim BytesNeeded As Long
    Dim PrinterName As String           ' Current printer
    Dim PrinterHandle As Long           ' Handle to printer
    Dim FormItem As String              ' For ListBox
    Dim RetVal As Long
    Dim FormSize As SIZEL               ' Size of desired form
    Dim FormName As String
    Dim ArrT
    
    PrinterName = Printer.DeviceName    ' Current printer
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        With FormSize   ' Desired page size
            .cx = 214000
            .cy = 216000
        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, _
                 NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, _
                 BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                ' List name and size including the count (index).
                FormName = Trim(PtrCtoVbString(.pName))
                FormItem = PtrCtoVbString(.pName) & " - " & .Size.cx / 1000 & _
                   " mm X " & .Size.cy / 1000 & " mm   (" & i + 1 & ")"
                
                If ValidateForms(FormName) = True Then
                    FormName = Trim(FormName)
                    FormName = Replace(FormName, " ", "", 1, Len(FormName))
                    FormName = UCase(FormName)
                    lstPrintForms.AddItem FormName
                    lstPrintForms.ItemData(lstPrintForms.NewIndex) = i + 1
                End If
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub


Private Sub Form_Resize()
    DoEvents
      btnFlatAll Me
    DoEvents
End Sub
