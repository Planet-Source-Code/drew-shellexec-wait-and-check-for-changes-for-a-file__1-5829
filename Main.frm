VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ShellExecute, Wait, and Check for changes."
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6240
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   15
      Top             =   945
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   855
      Left            =   405
      Picture         =   "Main.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   135
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "other file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   3
      Left            =   1380
      TabIndex        =   4
      Top             =   570
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".txt, .doc, .bmp, or some"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   2
      Left            =   3630
      TabIndex        =   3
      Top             =   330
      Width           =   2070
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "NOTE: This sample works on Windows 95, 98 and Windows NT."
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   1620
      TabIndex        =   2
      Top             =   1200
      Width           =   4605
   End
   Begin VB.Label Label1 
      Caption         =   "Click 'Open' and select a .txt, .doc, .bmp, or some other file to be opened by it's associated program."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   0
      Left            =   1380
      TabIndex        =   1
      Top             =   330
      Width           =   4485
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**(MODULE HEADER)**************(www.StotzerSoftware.com)**********
'*
'*   Module: frmMain
'*   Author: Andy Stotzer
'*    Email: Andy@Stotzer.com
'*      URL: http://www.StotzerSoftware.com
'*  Created: 01/01/1999
'*  Purpose: This Visual Basic Sample program demonstrates how to
'*           use the ShellExecute API call to load a file with it's
'*           associated program, wait for the program to end, then
'*           check the file's dates to see if the file has been
'*           changed.
'*
'*     Note: This code has been tested on Windows 95/98/NT.
'*
'******************************************************************

Private Const MAX_PATH = 260
Private Type FILETIME
        dwLowDateTime       As Long
        dwHighDateTime      As Long
End Type
Private Type SYSTEMTIME
        wYear               As Integer
        wMonth              As Integer
        wDayOfWeek          As Integer
        wDay                As Integer
        wHour               As Integer
        wMinute             As Integer
        wSecond             As Integer
        wMilliseconds       As Long
End Type
Private Type WIN32_FIND_DATA
        dwFileAttributes    As Long
        ftCreationTime      As FILETIME
        ftLastAccessTime    As FILETIME
        ftLastWriteTime     As FILETIME
        nFileSizeHigh       As Long
        nFileSizeLow        As Long
        dwReserved0         As Long
        dwReserved1         As Long
        cFileName           As String * MAX_PATH
        cAlternate          As String * 14
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

'**************************************************************
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const TH32CS_SNAPPROCESS As Long = 2&
'Private Const MAX_PATH           As Integer = 260
Private Const SW_SHOW           As Integer = 5
Private Type PROCESSENTRY32
    dwSize                      As Long
    cntUsage                    As Long
    th32ProcessID               As Long
    th32DefaultHeapID           As Long
    th32ModuleID                As Long
    cntThreads                  As Long
    th32ParentProcessID         As Long
    pcPriClassBase              As Long
    dwFlags                     As Long
    szExeFile                   As String * MAX_PATH
End Type
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Windows Version API Calls '
'**************************************************************
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long

Private Const PROCESS_VM_READ = 16
Private Type PROCESS_BASIC_INFORMATION
    ExitStatus                   As Long
    PebBaseAddress               As Long
    AffinityMask                 As Long
    BasePriority                 As Long
    UniqueProcessId              As Long
    InheritedFromUniqueProcessId As Long   'ParentProcessID'
End Type
Private Declare Function NtQueryInformationProcess _
                    Lib "ntdll" (ByVal ProcessHandle As Long, _
                                 ByVal ProcessInformationClass As Long, _
                                 ByRef ProcessInformation As PROCESS_BASIC_INFORMATION, _
                                 ByVal lProcessInformationLength As Long, _
                                 ByRef lReturnLength As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long



Private Sub cmdOpen_Click()

'**(PROCEDURE HEADER)***************(www.StotzerSoftware.com)******
'*
'*       Name: cmdOpen_Click
'*  Arguments: N/A
'*    Returns: NA/
'*     Author: Andy Stotzer
'*    Purpose: This Procedure allows the user to select the
'*             document/file, and call the ShellWaitCheck() function.
'*
'*  Developer        Date            Description
'*  ---------------------------------------------------------------
'*  Andy Stotzer     01/01/1999      Created Proc.
'*
'******************************************************************
    On Error GoTo ErrHandler
    
    'SHOW THE FILE OPEN DIALOGBOX, FOR THE USER TO SELECT A FILE'
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.Filter = "All Files (*.*)|*.*"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.ShowOpen
    
    'SHELL TO FILE, WAIT, CHECK FOR AND SAVE ANY CHANGES...
    If ShellWaitCheck(CommonDialog1.FileName) Then
        If vbYes = MsgBox("You have CHANGED this file." & vbLf & vbLf & "     '" & CommonDialog1.FileName & "'" & vbLf & vbLf & "Would you like to Save your changes to  ?", vbQuestion + vbYesNo, "You have changed this file!") Then
            'SAVE CHANGES'
            MsgBox "Do something with the file here.", vbInformation, "File Changed"
        End If
    Else
        MsgBox "The file you just edited, has not been changed.", vbInformation, "File not Changed"
    End If
    
Exit Sub
    
ErrHandler:
    'THE USER PRESSED THE CANCEL BUTTON'
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub


Private Function ShellWaitCheck(sFilename As String) As Boolean

'**(PROCEDURE HEADER)***************(www.StotzerSoftware.com)******
'*
'*       Name: ShellWaitCheck
'*  Arguments: (IN) Path & Filename to the file to be opened.
'*    Returns: True if file was changed.
'*             False if the file was NOT changed.
'*     Author: Andy Stotzer
'*    Purpose: This Procedure checks the files current 'Modified Date'.
'*             It then ShellExecutes and Waits for the program to end.
'*             Finally, it compares the files 'Modified Date' with the
'*             'Modified Date' saved before the ShellExecute.
'*
'*  Developer        Date            Description
'*  ---------------------------------------------------------------
'*  Andy Stotzer     01/01/1999      Created Proc.
'*
'******************************************************************
    Dim sOriginalCreateDate       As String
    Dim sOriginalLastModifiedDate As String
    Dim sNewCreateDate            As String
    Dim sNewLastModifiedDate      As String
    
    'GET DATES BEFORE SHELLING'
    Call GetFileDates(sFilename, sOriginalCreateDate, sOriginalLastModifiedDate)
    
    'SHELL OUT'
    Call ShellExecAndWait(sFilename, True, Me)
    
    'GET DATES AFTER SHELLING'
    Call GetFileDates(sFilename, sNewCreateDate, sNewLastModifiedDate)
    
    'COMPATE THE DATES FOR CHANGES'
    If sOriginalLastModifiedDate = "" Or sNewLastModifiedDate = "" Then
        ShellWaitCheck = False
    ElseIf DateDiff("s", sOriginalLastModifiedDate, sNewLastModifiedDate) > 0 Then
        ShellWaitCheck = True
    Else
        ShellWaitCheck = False
    End If
End Function


Public Sub GetFileDates(ByVal sFilename As String, _
                        ByRef sCreatedDate As String, _
                        ByRef sLastModifiedDate As String)
                        
'**(PROCEDURE HEADER)***************(www.StotzerSoftware.com)******
'*
'*       Name: GetFileDates
'*  Arguments: (IN)  - Path & Filename of the file to be checked.
'*             (OUT) - Variable to store the file's 'Created Date'.
'*             (OUT) - Variable to store the file's 'Last Modified Date'.
'*    Returns: n/a
'*     Author: Andy Stotzer
'*    Purpose: This Procedure fills in the passed variables with
'*             the files dates.
'*
'*  Developer        Date            Description
'*  ---------------------------------------------------------------
'*  Andy Stotzer     01/01/1999      Created Proc.
'*
'******************************************************************
    Dim hFile   As Long
    Dim WFD     As WIN32_FIND_DATA
    Dim ST      As SYSTEMTIME
    Dim lRC     As Long
    Dim ds      As Single
    
    hFile = FindFirstFile(sFilename, WFD)
    If hFile > 0 Then
        'CONVERT THE CREATION TIME'
        lRC = FileTimeToSystemTime(WFD.ftCreationTime, ST)
        If lRC Then
            sCreatedDate = Format(ST.wYear & "/" & ST.wMonth & "/" & ST.wDay & " " & _
                                  ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond, _
                                  "mm/dd/yy hh:mm:ss")
        End If
        
        'CONVERT THE LAST WRITE TIME'
        lRC = FileTimeToSystemTime(WFD.ftLastWriteTime, ST)
        If lRC Then
            sLastModifiedDate = Format(ST.wYear & "/" & ST.wMonth & "/" & ST.wDay & " " & _
                                  ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond, _
                                  "mm/dd/yy hh:mm:ss")
        End If
    Else
        sCreatedDate = "Error"
        sLastModifiedDate = "Error"
    End If
End Sub

Public Sub ShellExecAndWait(sFilename As String, _
                            bWaitToEnd As Boolean, _
                            Optional oForm As Variant)
                            
'**(PROCEDURE HEADER)***************(www.StotzerSoftware.com)******
'*
'*       Name: ShellExecAndWait
'*  Arguments: (IN) - Path & Filename of the file to be checked.
'*             (IN) - Flag to specify whether to wait or not.
'*             (IN) - Optional - Form object.
'*    Returns: N/A
'*     Author: Andy Stotzer
'*    Purpose: This Procedure uses the ShellExecute API call to
'*             load a file into it's associated program, and wait
'*             for the program to end.
'*
'*             First, this procedure gets it's own ProcessID.  Then
'*             it calls ShellExecute() to load the file.  Next, it
'*             checks if it should wait or not.  If it should wait,
'*             it checks to see if it is running on Windows 9x or NT.
'*             It then calls the appropriate API calls for the OS
'*             Platform to obtain the first Process that has this
'*             process as a parent.  And last, it waits for that
'*             Process to end.
'*
'*  -- WARNING: ---------------------------------------------------
'*  This procedure may/will NOT WORK correctly if you Shell to more
'*  than one application at a time.  It only looks for the first
'*  Process that has your program as the Parent.
'*  ---------------------------------------------------------------
'*
'*  Developer        Date            Description
'*  ---------------------------------------------------------------
'*  Andy Stotzer     01/01/1999      Created Proc.
'*
'******************************************************************
                            
    Dim hSnapShot           As Long
    Dim hProcess            As Long
    Dim uProcess            As PROCESSENTRY32
    Dim lRC                 As Long
    Dim lShelledProcessID   As Long
    Dim lMyProcessID        As Long
    Dim myOS                As OSVERSIONINFO
    Dim iWinVersion         As Integer
    
    '--- GET THIS APPS PROCESS ID ---'
    lMyProcessID = GetCurrentProcessId
    
    
    '--- SHELLEXECUTE THE DOCUMENT TO THE ASSOCIATED APP ---'
    If IsMissing(oForm) Then
        lRC = ShellExecute(0&, "open", sFilename, vbNullString, CurDir$, SW_SHOW)
    Else
        lRC = ShellExecute(oForm.hwnd, "open", sFilename, vbNullString, CurDir$, SW_SHOW)
    End If
    If lRC < 32 Then
        MsgBox "Unable to open this file: " & vbLf & " " & sFilename, vbExclamation, "Error Opening File"
        Exit Sub
    End If

    '-----------------------'
    'Wait for Process to End'
    '-----------------------'
    If Not bWaitToEnd Then
        Exit Sub
    End If
    
    
    'FIND OUT WHAT VERSION OF WINDOWS IS RUNNING'
    myOS.dwOSVersionInfoSize = Len(myOS)
    GetVersionEx myOS
    If myOS.dwMajorVersion = 4 Then
        iWinVersion = myOS.dwPlatformId
    Else
        MsgBox "You need to upgrade your version of windows to use the 'Wait for Process to End' feature of this function: 'ShellExecAndWait'.", vbExclamation, "Sorry!"
        Exit Sub
    End If
    
    
    'WHAT VERSION OF WINDOWS ARE THEY RUNNING?'
    Select Case iWinVersion
        Case VER_PLATFORM_WIN32_NT
            Dim cb                  As Long
            Dim cbNeeded            As Long
            Dim NumElements         As Long
            Dim ProcessIDs()        As Long
            Dim lRet                As Long
            Dim i                   As Long
            
            'GET AN ARRAY OF PROCESS IDS'
            cb = 8
            cbNeeded = 96
            Do While cb <= cbNeeded
                cb = cb * 2
                ReDim ProcessIDs(cb / 4) As Long
                lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
            Loop
            NumElements = cbNeeded / 4
            
            For i = 1 To NumElements
                'Get a handle to the Process
                hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessIDs(i))
                
                '--------------------------------'
                '  FIND THE NT PARENT PROCESS ID '
                '--------------------------------'
                Dim lntStatus                   As Long
                Dim lProcessHandle              As Long
                Dim lProcessBasicInfo           As Long
                Dim tProcessInformation         As PROCESS_BASIC_INFORMATION
                Dim lProcessInformationLength   As Long
                Dim lReturnLength               As Long
                
                'INITIALIZE EVERYTHING'
                lProcessHandle = hProcess
                lProcessBasicInfo = 0               '0 = REQUEST PROCESS INFORMATION'
                lProcessInformationLength = Len(tProcessInformation)
                
                'CALL NT FUNCTION TO FIND PROCESS INFORMATION'
                lntStatus = NtQueryInformationProcess(lProcessHandle, lProcessBasicInfo, tProcessInformation, lProcessInformationLength, lReturnLength)
                
                'AM I THE PARENT OF THIS PROCESS?'
                If lMyProcessID = tProcessInformation.InheritedFromUniqueProcessId Then
                    lShelledProcessID = tProcessInformation.UniqueProcessId
                    'CLOSE THE HANDLE TO THE PROCESS'
                    CloseHandle (hProcess)
                    Exit For
                End If
                
                'CLOSE THE HANDLE TO THE PROCESS'
                CloseHandle (hProcess)
            Next
            
            
        Case VER_PLATFORM_WIN32_WINDOWS
            ' WINDOWS 95/98 '
            '--------------------------------'
            '  FIND THE NT PARENT PROCESS ID '
            '--------------------------------'
            hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
            If hSnapShot = 0 Then Exit Sub
            uProcess.dwSize = Len(uProcess)
            lRC = ProcessFirst(hSnapShot, uProcess)
            Do While lRC
                'AM I THE PARENT OF THIS PROCESS?'
                If lMyProcessID = uProcess.th32ParentProcessID Then
                    'FOUND IT'
                    lShelledProcessID = uProcess.th32ProcessID
                    Exit Do
                End If
                lRC = ProcessNext(hSnapShot, uProcess)
            Loop
            'CLOSE THE HANDLE TO THE PROCESS'
            Call CloseHandle(hSnapShot)
        
        Case Else
        
    End Select
    

    '--- WAIT FOR PROCESS TO END ---'
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, lShelledProcessID)
    Do
        Call GetExitCodeProcess(hProcess, lRC)
        DoEvents
    Loop While lRC > 0

    Call CloseHandle(hProcess)
        
End Sub

