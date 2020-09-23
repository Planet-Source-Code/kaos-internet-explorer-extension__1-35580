VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Internet Explorer Extension"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUninstall 
      Caption         =   "UnInstall"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtInformation 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmMain.frx":0E42
      Top             =   120
      Width           =   5295
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Install"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   480
      Picture         =   "frmMain.frx":0FEC
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const GUID As String = "{4B3520B0-D518-4443-BA9E-2D4CE7F773C5}"
Const ADDRESS As String = "Software\Microsoft\Internet Explorer\Extensions\" & GUID


Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdInstall_Click()
        
    'Dim fso As Object
    Dim fso As New FileSystemObject
    Dim sErrMsg As String
    Set fso = CreateObject("scripting.filesystemobject")
    
    If fso.FileExists("launch.exe") = False Then
        sErrMsg = "launch.exe is not found in the current directory."
        Me.txtInformation = Me.txtInformation & vbCrLf & sErrMsg
        MsgBox sErrMsg, vbCritical
        Exit Sub
    End If
    
    If fso.FileExists("HotIcon.ico") = False Then
        sErrMsg = "HotIcon.ico is not found in the current directory. Please copy it in the current directory and try again."
        Me.txtInformation = Me.txtInformation & vbCrLf & sErrMsg
        MsgBox sErrMsg, vbCritical
        Exit Sub
    End If
    
    If fso.FileExists("Icon.ico") = False Then
        sErrMsg = "Icon.ico is not found in the current directory. Please copy it in the current directory and try again."
        Me.txtInformation = Me.txtInformation & vbCrLf & sErrMsg
        MsgBox sErrMsg, vbCritical
        Exit Sub
    End If
    
    
    
    
    'find if the guid is already registered
    If IsInstalled() = False Then
        InsatllExtension
        
    Else
        Dim ans As VbMsgBoxResult
        ans = MsgBox("This extension already exists in your Internet Explorer. Do you want to run it again?", vbQuestion & vbYesNo)
        If ans = vbYes Then
            InsatllExtension
        End If
    End If
End Sub


' checks if the extension is already installed in this machine
Private Function IsInstalled() As Boolean
    On Error GoTo lblErrHandler:
        
        Dim lngRegKey As Long
        
        Call RegOpenKey(HKEY_LOCAL_MACHINE, ADDRESS, lngRegKey)
            'MsgBox lngRegKey
            If lngRegKey = 0 Then
                ' There is no key of this name
                ' this extension is not installed
                IsInstalled = False
            Else
                ' the extension is already installed
                IsInstalled = True
            End If
        Call RegCloseKey(lngRegKey)
    Exit Function
lblErrHandler:
    ToString
End Function



' installs the extensions
Private Function InsatllExtension() As Integer
    On Error GoTo lblErrHandler:
        
        Dim hNewKey As Long         'handle to the new key
        Dim lRetVal As Long         'result of the RegCreateKeyEx function
        
        ' Create the key with the guid
        lRetVal = RegCreateKeyEx(HKEY_LOCAL_MACHINE, ADDRESS, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
        RegCloseKey (hNewKey)
        
       'Create and populate ButtonText
       ' REG_SZ
       ' CMM Automation
       lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, ADDRESS, 0, KEY_ALL_ACCESS, hNewKey)
       lRetVal = SetValueEx(hNewKey, "ButtonText", REG_SZ, "Extension")
       RegCloseKey (hNewKey)


       'Create and populate CLSID
       ' REG_SZ
       ' {1FBA04EE-3024-11d2-8F1F-0000F87ABD16}
       lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, ADDRESS, 0, KEY_ALL_ACCESS, hNewKey)
       lRetVal = SetValueEx(hNewKey, "CLSID", REG_SZ, "{1FBA04EE-3024-11d2-8F1F-0000F87ABD16}")
       RegCloseKey (hNewKey)
       
       
       'Create and populate Default Visible
       ' REG_SZ
       ' yes
       lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, ADDRESS, 0, KEY_ALL_ACCESS, hNewKey)
       lRetVal = SetValueEx(hNewKey, "Default Visible", REG_SZ, "Yes")
       RegCloseKey (hNewKey)
       
       'Create and populate Default Exec
       ' REG_SZ
       ' yes
       lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, ADDRESS, 0, KEY_ALL_ACCESS, hNewKey)
       lRetVal = SetValueEx(hNewKey, "Exec", REG_SZ, App.Path & "\launch.exe")
       RegCloseKey (hNewKey)
       
       

       'Create and populate HotIcon
       ' REG_SZ
       ' yes
       lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, ADDRESS, 0, KEY_ALL_ACCESS, hNewKey)
       lRetVal = SetValueEx(hNewKey, "HotIcon", REG_SZ, App.Path & "\HotIcon.ico")
       RegCloseKey (hNewKey)
       

       'Create and populate Icon
       ' REG_SZ
       ' yes
       lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, ADDRESS, 0, KEY_ALL_ACCESS, hNewKey)
       lRetVal = SetValueEx(hNewKey, "Icon", REG_SZ, App.Path & "\Icon.ico")
       RegCloseKey (hNewKey)


       'Create and populate MenuSatusBar
       ' REG_SZ
       ' yes
       lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, ADDRESS, 0, KEY_ALL_ACCESS, hNewKey)
       lRetVal = SetValueEx(hNewKey, "MenuStatusBar", REG_SZ, "Opens the Application.")
       RegCloseKey (hNewKey)


       'Create and populate MenuText
       ' REG_SZ
       ' yes
       lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, ADDRESS, 0, KEY_ALL_ACCESS, hNewKey)
       lRetVal = SetValueEx(hNewKey, "MenuText", REG_SZ, "Extension")
       RegCloseKey (hNewKey)
       
       Dim sErrMsg As String
        sErrMsg = "Internet Explorer extensions have been installed successfully in your machine. Open a new instance of IE to see the changes."
        Me.txtInformation = Me.txtInformation & sErrMsg
        MsgBox sErrMsg, vbInformation
       

    Exit Function
lblErrHandler:
    ToString
End Function


' checks if the file is present in the current directory
Public Function IsPresent(sFileName As String) As Boolean
    On Error GoTo lblErrHandler:
        
        sFileName = App.Path & "\" & sFileName
        ' connect to the file and create a textstream
        Dim fso As Object
        'Dim fso As New FileSystemObject
        Set fso = CreateObject("scripting.filesystemobject")
        fso.FileExists (sFileName)
        
        
    Exit Function
lblErrHandler:
    ToString

End Function

' delete the key
Private Sub cmdUninstall_Click()
    
    'find if the guid is already registered
    If IsInstalled() = True Then
        UnInstall
    Else
        MsgBox "This extension does not exis in your Internet Explorer.", vbInformation
    End If

End Sub


' Uninstall
Public Sub UnInstall()
    On Error GoTo lblErrHandler:
        
        Dim hNewKey As Long         'handle to the new key
        Dim lRetVal As Long         'result of the RegCreateKeyEx function
        
        ' Create the key with the guid
        lRetVal = DeleteKey(HKEY_LOCAL_MACHINE, ADDRESS)
        
        MsgBox "Extension has been successfully uninstalled. Open a new instance of Internet Explorer to see the changes.", vbInformation
        
            
    Exit Sub
lblErrHandler:
    ToString

End Sub

