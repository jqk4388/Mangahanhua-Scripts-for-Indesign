' ============================================
' Manga Layout Automation - VBScript Launcher
' Manga Layout Automation - VBScript Launcher
' 
' Features:
' - Launch InDesign application
' - Call manga_layout.jsx script
' - Handle error output and delay management
' 
' Version: 1.0.0
' ============================================

Option Explicit

' Constant definitions
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

' Global variables
Dim app, fso, scriptPath, logPath
Dim startTime, endTime, duration

' Initialization
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)
logPath = scriptPath & "\manga_layout_vbs.log"

' Record start time
startTime = Timer

' Main function
Sub Main()
    On Error Resume Next
    
    ' Write to log
    Call LogMessage("===== Manga Layout VBS Launcher Started =====")
    Call LogMessage("Script path: " & scriptPath)
    Call LogMessage("Start time: " & Now())
    
    ' Check if InDesign is already running
    Dim isAlreadyRunning
    isAlreadyRunning = IsInDesignRunning()
    
    ' Get InDesign application object
    Call LogMessage("Connecting to InDesign application...")
    Set app = GetInDesignApplication()
    
    If Err.Number <> 0 Then
        Call LogMessage("Error: Unable to connect to InDesign - " & Err.Description)
        Call LogMessage("Please ensure InDesign is installed and running")
        WScript.Quit(1)
    End If
    
    Call LogMessage("Successfully connected to InDesign")
    
    ' Delay to wait for InDesign to fully load
    Call Delay(1000)
    
    ' Set script file path
    Dim jsxPath
    jsxPath = scriptPath & "\manga_layout.jsx"
    
    ' Check if script file exists
    If Not fso.FileExists(jsxPath) Then
        Call LogMessage("Error: Script file does not exist - " & jsxPath)
        WScript.Quit(1)
    End If
    
    Call LogMessage("Script file: " & jsxPath)
    
    ' Check if config file exists
    Dim configPath
    configPath = scriptPath & "\manga_layout_config.json"
    
    If Not fso.FileExists(configPath) Then
        Call LogMessage("Warning: Config file does not exist - " & configPath)
        Call LogMessage("Will run with default configuration")
    Else
        Call LogMessage("Config file: " & configPath)
    End If
    
    ' Execute JSX script
    Call LogMessage("Starting JSX script execution...")
    
    ' Execute using DoScript and capture errors
    Dim result
    result = ExecuteJSX(jsxPath)
    
    ' Record end time
    endTime = Timer
    duration = endTime - startTime
    
    ' Process results
    If Err.Number <> 0 Then
        Call LogMessage("Error: Script execution failed - " & Err.Description)
        Call LogMessage("Error code: " & Err.Number)
        Call LogMessage("Execution duration: " & FormatNumber(duration, 2) & " seconds")
        WScript.Quit(1)
    Else
        Call LogMessage("Script execution completed")
        Call LogMessage("Execution duration: " & FormatNumber(duration, 2) & " seconds")
        Call LogMessage("===== Execution successful =====")
    End If
    
    ' Cleanup
    Set app = Nothing
    Set fso = Nothing
    
    WScript.Quit(0)
End Sub

' Get InDesign application object
Function GetInDesignApplication()
    On Error Resume Next
    
    Dim indApp
    Dim version, versions
    versions = Array("2025", "2024", "2023", "2022", "2021", "2020", "CC 2019", "CC 2018")
    
    ' Try to get running instance
    For Each version In versions
        Set indApp = Nothing
        Err.Clear
        
        On Error Resume Next
        Set indApp = GetObject("", "InDesign.Application." & version)
        On Error Resume Next
        
        If Not indApp Is Nothing Then
            Call LogMessage("Connected to InDesign " & version)
            Set GetInDesignApplication = indApp
            Exit Function
        End If
    Next
    
    ' Try to create new instance
    For Each version In versions
        Set indApp = Nothing
        Err.Clear
        
        On Error Resume Next
        Set indApp = CreateObject("InDesign.Application." & version)
        On Error Resume Next
        
        If Not indApp Is Nothing Then
            Call LogMessage("Created InDesign " & version & " instance")
            Set GetInDesignApplication = indApp
            Exit Function
        End If
    Next
    
    ' If all fail, try without version number
    Err.Clear
    On Error Resume Next
    Set indApp = GetObject("", "InDesign.Application")
    On Error Resume Next
    
    If Not indApp Is Nothing Then
        Call LogMessage("Connected to InDesign (default version)")
        Set GetInDesignApplication = indApp
        Exit Function
    End If
    
    ' Last attempt to create
    Err.Clear
    On Error Resume Next
    Set indApp = CreateObject("InDesign.Application")
    On Error Resume Next
    
    Set GetInDesignApplication = indApp
End Function

' Check if InDesign is running
Function IsInDesignRunning()
    On Error Resume Next
    
    Dim wmi, processes, process
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set processes = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name LIKE '%InDesign%'")
    
    IsInDesignRunning = (processes.Count > 0)
End Function

' Execute JSX script
Function ExecuteJSX(jsxPath)
    On Error Resume Next
    
    ' Script language constant
    Const idJavascript = 1246973031
    
    ' Execute script
    app.DoScript jsxPath, idJavascript
    
    If Err.Number <> 0 Then
        ExecuteJSX = False
        Call LogMessage("DoScript error: " & Err.Description)
    Else
        ExecuteJSX = True
    End If
End Function

' Delay function
Sub Delay(milliseconds)
    Dim startTime, currentTime
    startTime = Timer
    
    Do
        currentTime = Timer
        ' Handle Windows messages to prevent the program from appearing frozen
        WScript.Sleep 100
    Loop While (currentTime - startTime) * 1000 < milliseconds
End Sub

' Log message function
Sub LogMessage(message)
    On Error Resume Next
    
    Dim logFile
    Set logFile = fso.OpenTextFile(logPath, ForAppending, True)
    
    logFile.WriteLine Now() & " - " & message
    logFile.Close
    
    ' Also output to console (for debugging)
    WScript.Echo message
End Sub

' Execute main function
Main
