Attribute VB_Name = "modConnection"
Public CON As New ADODB.Connection
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000


Sub main()
         On Error GoTo errConnection
        If CON.State = 1 Then ' If Connection is already Opened
        
        Else
              Screen.MousePointer = vbHourglass
              CON.ConnectionTimeout = 500
             CON.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\srx\Desktop\ReportOptions\dbReportOptions.mdb;Persist Security Info=False")
             CON.CursorLocation = adUseClient
             Screen.MousePointer = vbDefault
             Load frmMain
             frmMain.Show
          End If

        Exit Sub
errConnection:
    MsgBox "DataBase Server Not Available " & Chr(13) & "Please Contact IT Department", vbInformation, "FACOR ERP"
    End
End Sub

Public Function PopulateComboBoxWithDefinedIndex(combo As ComboBox, LrsTemp As ADODB.Recordset, Optional IDtoPOPULATE As Integer = 0, Optional ValueToPopulate As Integer = 1, Optional ClearVal As Boolean, Optional NoDuplicate As Boolean = False)
'Function to Populate desired Values of any Recordset in ComboBox
' This Functions Automatically and Essentialy Sets the RecordSet's "0 th" Values in ItemData
'Created by Nitesh on February 28, 2004
        Dim Cnt As Integer
          On Error GoTo ErrPopulateTextBox
                If ClearVal = True Then
                    combo.Clear
                End If
            If LrsTemp.EOF = True Or LrsTemp.BOF = True Then
            Else
            
                'If NoDuplicate = False Then
                    While Not LrsTemp.EOF = True
                        combo.AddItem Trim(LrsTemp.Fields(ValueToPopulate).Value)
                        combo.ItemData(combo.NewIndex) = LrsTemp.Fields(IDtoPOPULATE).Value
                        LrsTemp.MoveNext
                    Wend
                 'End If
                 
'                 If NoDuplicate = True Then
'                    While Not LrsTemp.EOF = True
'                        For Cnt = 0 To combo.ListCount - 1
'                            If LrsTemp.Fields(rsindex).Value = combo.List(Cnt) Then
'                                LrsTemp.MoveNext
'                            Else
'                                combo.AddItem Trim(LrsTemp.Fields(rsindex).Value)
'                                combo.ItemData(combo.ListCount - 1) = LrsTemp.Fields(0).Value
'                                LrsTemp.MoveNext
'                            End If
'                         Next Cnt
'                     Wend
'                 End If
              ' combo.Text = combo.List(0)
              combo.ListIndex = 0
             End If
             
        
        Exit Function
ErrPopulateTextBox:
        MsgBox Err.Description
        Exit Function
End Function

Public Function btnFlatAll(Container As Form)
Dim Button
Dim Value As Boolean
For Each Button In Container.Controls
    If TypeOf Button Is CommandButton Or TypeOf Button Is Frame Then
        
        SetWindowLong Button.hwnd, GWL_STYLE, WS_CHILD Or BS_FLAT
        Button.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
    End If
Next
End Function
