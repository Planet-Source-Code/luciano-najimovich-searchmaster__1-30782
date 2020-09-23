Attribute VB_Name = "mMain"
Option Explicit
Option Base 1

'First, I defined the global constants for the application
Global Const APP_Name = "SearchMaster"
Global Const APP_Version = "1.0"

Function GetCommandLine(CmdLine As String, Optional MaxArgs As Integer) As Variant
   'Declare variables.
   Dim C, CmdLnLen, InArg, I, NumArgs
   'See if MaxArgs was provided.
   If IsMissing(MaxArgs) Then MaxArgs = 10
   'Make array of the correct size.
   ReDim ArgArray(MaxArgs)
   NumArgs = 0: InArg = False
   CmdLnLen = Len(CmdLine)
   'Go thru command line one character
   'at a time.
   For I = 1 To CmdLnLen
      C = Mid(CmdLine, I, 1)
      'Test for space or tab.
      If (C <> "-" And C <> vbTab And C <> "  ") Then
         'Neither space nor tab.
         'Test if already in argument.
         If Not InArg Then
         'New argument begins.
         'Test for too many arguments.
            If NumArgs = MaxArgs Then Exit For
            NumArgs = NumArgs + 1
            InArg = True
         End If
         'Concatenate character to current argument.
         ArgArray(NumArgs) = ArgArray(NumArgs) & C
      Else
         'Found a space or tab.
         'Set InArg flag to False.
         InArg = False
      End If
   Next I
   'Resize array just enough to hold arguments.
   ReDim Preserve ArgArray(NumArgs)
   'Return Array in Function name.
   GetCommandLine = ArgArray()
End Function

Function LoadFile(CDialog As CommonDialog, Optional InitPath As String, Optional Filter As String, Optional Title As String) As String
 
 If Not IsMissing(InitPath) Then
   CDialog.InitDir = InitPath
 End If
 
 If Not IsMissing(Filter) Then
   CDialog.Filter = Filter
 End If
 
 If Not IsMissing(Title) Then
   CDialog.Filter = Filter
 End If
  
 CDialog.Filename = ""
 CDialog.ShowOpen
 LoadFile = CDialog.Filename
 
End Function

Function VerifyFileOk(Filename As String) As Boolean
 
 VerifyFileOk = True
 
 If Filename = "" Then
   MsgBox "Please select an input filename!", vbCritical + vbOKOnly, APP_Name
   VerifyFileOk = False
 ElseIf Dir(Filename) = "" Then
   MsgBox "The filename doesn't exist. Please enter a valid filename.", vbCritical + vbOKOnly, APP_Name
   VerifyFileOk = False
 End If
  
End Function

Function CheckAllOk(FileIn As String, FileOut As String, SearchString As String, ReplaceString As String) As Boolean
  
  CheckAllOk = True

  If Not VerifyFileOk(FileIn) Then
     CheckAllOk = False
  ElseIf FileOut = "" Then
    MsgBox "Please select an output filename!", vbCritical + vbOKOnly, APP_Name
    CheckAllOk = False
  ElseIf SearchString = "" Then
    MsgBox "Please select a search string!", vbCritical + vbOKOnly, APP_Name
    CheckAllOk = False
  ElseIf ReplaceString = "" Then
    MsgBox "Please select a replace string!", vbCritical + vbOKOnly, APP_Name
    CheckAllOk = False
  End If

End Function
Sub SearchAndReplace(FileIn As String, FileOut As String, SearchString As String, ReplaceString As String, EmbeddedText As Boolean, Status As ProgressBar)
Dim Response As Integer
Dim CheckOK As Boolean
Dim oFileSysObj As FileSystemObject
Dim oInFile As TextStream
Dim oOutFile As TextStream
Dim AuxLine As String
Dim Max As Integer

Dim I, j As Long


  On Error GoTo Err_Trap
  
   FileIn = Trim(FileIn)
   FileOut = Trim(FileOut)
   SearchString = Trim(SearchString)
   ReplaceString = Trim(ReplaceString)
   
   If Dir(FileOut) <> "" Then
     
     Response = MsgBox("The output file already exist. You want to overwrite it?", vbExclamation + vbYesNo, APP_Name)
     If Response = vbYes Then
       Kill (FileOut)
     Else
       Exit Sub
     End If
   
   End If
   
   Max = 0
   
   Set oFileSysObj = New FileSystemObject
   Set oInFile = oFileSysObj.OpenTextFile(FileIn, ForReading, False, TristateUseDefault)
   Set oOutFile = oFileSysObj.CreateTextFile(FileOut, True)
      
   If oInFile.AtEndOfStream Then
      MsgBox "The input file is empty! Please select another file.", vbExclamation + vbOKOnly, APP_Name
      Exit Sub
   End If
   
   While Not oInFile.AtEndOfStream
         AuxLine = oInFile.ReadLine
         Max = Max + 1
   Wend
   
   oInFile.Close
   Set oInFile = oFileSysObj.OpenTextFile(FileIn, ForReading, False, TristateUseDefault)
   
   Status.Min = 0
   Status.Max = Max
   Status.Value = 0
   
   
   If Not EmbeddedText Then
      SearchString = Space(1) & SearchString & Space(1)
      ReplaceString = Space(1) & ReplaceString & Space(1)
   End If
   
   While Not oInFile.AtEndOfStream
      AuxLine = oInFile.ReadLine
      AuxLine = Replace(AuxLine, SearchString, ReplaceString, , , vbTextCompare)
      oOutFile.WriteLine AuxLine
      Status.Value = Status.Value + 1
   Wend
   
   oInFile.Close
   
   Set oFileSysObj = Nothing
   
   MsgBox "The output file was created sucessfully!", vbInformation + vbOKOnly, APP_Name
   
  Exit Sub
  
Err_Trap:
   Set oFileSysObj = Nothing
   oInFile.Close
   MsgBox "An error was occurred. Error code " & Err.Number & " -" & Err.Description, vbCritical + vbOKOnly, "Oops!"
End Sub

Sub Main()
Dim I As Byte
Dim Cant As Byte
Dim arrParam As Variant
Dim MsgErrorParam As String
Dim Aux As Boolean

 On Error GoTo Err_Trap
   
  If Command() <> "" Then
   
   arrParam = GetCommandLine(Command(), 5)
   Cant = UBound(arrParam)
   
   If Cant >= 1 And Cant > 3 Then
     
     If CheckParams(arrParam) Then
      
      Frm_Main.Show
      Frm_Main.txtInput.Text = Right(arrParam(1), Len(arrParam(1)) - 2)
      Frm_Main.txtOutput.Text = Right(arrParam(2), Len(arrParam(2)) - 2)
      Frm_Main.txtSearch.Text = Right(arrParam(4), Len(arrParam(4)) - 2)
      Frm_Main.txtReplace.Text = Right(arrParam(5), Len(arrParam(5)) - 2)
      
      If Cant > 4 Then
          If UCase(Right(arrParam(3), Len(arrParam(3)) - 2)) = "TRUE" Then Frm_Main.chkEmbedded.Value = 1 ' Checked!
      End If
      
      If CheckAllOk(Frm_Main.txtInput, Frm_Main.txtOutput, Frm_Main.txtSearch, Frm_Main.txtReplace) Then
        Frm_Main.lblStatusMessage.Visible = True
        Frm_Main.pgbStatus.Visible = True
        Call SearchAndReplace(Frm_Main.txtInput, Frm_Main.txtOutput, Frm_Main.txtSearch, Frm_Main.txtReplace, Frm_Main.chkEmbedded, Frm_Main.pgbStatus)
        Frm_Main.lblStatusMessage.Visible = False
        Frm_Main.pgbStatus.Visible = False
      End If
     
     End If
   
   Else
     Aux = CheckParams(arrParam)
   End If
  
  End If
  
  Frm_Main.Show
  
  Exit Sub

Err_Trap:
   MsgBox Err.Number & " " & Err.Description
End Sub
Function CheckParams(Vec As Variant) As Boolean
Dim MsgErrorParam As String

  CheckParams = True
  
  If UBound(Vec) < 4 Then
   MsgErrorParam = " Parameters for command line input :" & vbCrLf & vbCrLf
   MsgErrorParam = MsgErrorParam & "-i  [filename]      : Filename (with path) for the input file." & vbCrLf
   MsgErrorParam = MsgErrorParam & "-o [filename]       : Filename (with path) for the output file." & vbCrLf
   MsgErrorParam = MsgErrorParam & "-s [string]         : String to search in the input filename." & vbCrLf
   MsgErrorParam = MsgErrorParam & "-r [string]         : String to use to replace the search string." & vbCrLf
   MsgErrorParam = MsgErrorParam & "-e [true][false]    : Use the option to searcg in the sdasdas or not. Default is False." & vbCrLf
   MsgBox MsgErrorParam, vbInformation + vbOKOnly, APP_Name
   CheckParams = False
  End If
  
End Function
