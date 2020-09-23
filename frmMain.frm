VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Platform of any eVB Project"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dialog 
      Left            =   3600
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.ebp | *.ebp"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   615
      Left            =   1200
      TabIndex        =   4
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Choose Platform"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4335
      Begin VB.CommandButton cmdSet 
         Caption         =   "Set Platform"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cmbPlat 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose eVB Project"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse ..."
         Height          =   615
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtPath 
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************
' App for change the platform of a eVB project
' Author: Toni Novak 28 / 08 / 2002
'*****************************************************

Private Type PlatType
    PlatGUID As String
    DevGUID As String
    ProjType As String
    Platform As String
End Type

Private Const platHPC = 0
Private Const platPalmSize = 1
Private Const platPC = 2
Private Const platPC2002 = 3

Dim PlatType(0 To 4) As PlatType
Private Sub cmdBrowse_Click()

dialog.ShowOpen

If dialog.FileName <> "" Then

    txtPath.Text = dialog.FileName
    cmbPlat.Enabled = True
    cmdSet.Enabled = True
    
End If

End Sub
Sub SetPlat()

On Error GoTo ErrTrap

If dialog.FileName <> "" Then

    Dim FByte As Byte
    Dim cLine As String
    Dim OutText As String
    Open dialog.FileName For Binary As #1
        Do
            Get #1, , FByte
            If FByte <> 13 And FByte <> 10 Then cLine = cLine & Chr(FByte)
            If FByte = 13 Then
                
                If LCase(Left(cLine, Len("platformguid="))) = "platformguid=" Then
                
                    Select Case cmbPlat.ListIndex
                        Case platHPC
                            cLine = "PlatformGUID=" & PlatType(platHPC).PlatGUID
                        Case platPalmSize
                            cLine = "PlatformGUID=" & PlatType(platPalmSize).PlatGUID
                        Case platPC
                            cLine = "PlatformGUID=" & PlatType(platPC).PlatGUID
                        Case platPC2002
                            cLine = "PlatformGUID=" & PlatType(platPC2002).PlatGUID
                    End Select
                
                End If
                
                If LCase(Left(cLine, Len("deviceguid="))) = "deviceguid=" Then
                
                    Select Case cmbPlat.ListIndex
                        Case platHPC
                            cLine = "DeviceGUID=" & PlatType(platHPC).DevGUID
                        Case platPalmSize
                            cLine = "DeviceGUID=" & PlatType(platPalmSize).DevGUID
                        Case platPC
                            cLine = "DeviceGUID=" & PlatType(platPC).DevGUID
                        Case platPC2002
                            cLine = "DeviceGUID=" & PlatType(platPC2002).DevGUID
                    End Select
                
                
                End If
                
                If LCase(Left(cLine, Len("projecttype="))) = "projecttype=" Then
                
                    Select Case cmbPlat.ListIndex
                        Case platHPC
                            cLine = "ProjectType=" & PlatType(platHPC).ProjType
                        Case platPalmSize
                            cLine = "ProjectType=" & PlatType(platPalmSize).ProjType
                        Case platPC
                            cLine = "ProjectType=" & PlatType(platPC).ProjType
                        Case platPC2002
                            cLine = "ProjectType=" & PlatType(platPC2002).ProjType
                    End Select
                
                End If
                                
                If LCase(Left(cLine, Len("platform="))) = "platform=" Then
                
                    Select Case cmbPlat.ListIndex
                        Case platHPC
                            cLine = "Platform=" & PlatType(platHPC).Platform
                        Case platPalmSize
                            cLine = "Platform=" & PlatType(platPalmSize).Platform
                        Case platPC
                            cLine = "Platform=" & PlatType(platPC).Platform
                        Case platPC2002
                            cLine = "Platform=" & PlatType(platPC2002).Platform
                    End Select
                                
                End If
                
                OutText = OutText & cLine & vbCrLf
                
                'Debug.Print cLine
                cLine = ""
                                
            End If
        Loop Until EOF(1)
    Close #1
    
    Open dialog.FileName For Output As #1
        Print #1, OutText
    Close #1
    
    MsgBox "Platform changed !!", vbInformation, "eVB Platform"
    
End If

Exit Sub

ErrTrap:

MsgBox "Error ! No change was made: " & vbCrLf & Err.Description, vbCritical, "eVB Platform"

End Sub
Private Sub cmdSet_Click()

SetPlat

End Sub
Private Sub Command1_Click()

End

End Sub
Private Sub Form_Load()

'Load GUIDs definitions

PlatType(platHPC).PlatGUID = "{74239C21-1DCA-11D2-9747-00A0240918F0}"
PlatType(platHPC).DevGUID = "{6CEF7360-4355-11D2-975C-00A0240918F0}"
PlatType(platHPC).ProjType = ""
PlatType(platHPC).Platform = ""

PlatType(platPalmSize).PlatGUID = "{458BFDB0-A6A6-11D2-BBCF-00A0C9C9CCEE}"
PlatType(platPalmSize).DevGUID = "{6CEF7360-4355-11D2-975C-00A0240918F0}"
PlatType(platPalmSize).ProjType = ""
PlatType(platPalmSize).Platform = ""

PlatType(platPC).PlatGUID = "{6D5C6210-E14B-11D2-B72A-0000F8026CEE}"
PlatType(platPC).DevGUID = "{3CFA6F81-EB79-11D2-BAC5-006097BA8DF0}"
PlatType(platPC).ProjType = "WinCE"
PlatType(platPC).Platform = "{6D5C6210-E14B-11D2-B72A-0000F8026CEE}"

PlatType(platPC2002).PlatGUID = "{DE9660AC-85D3-4C63-A6AF-46A3B3B83737}"
PlatType(platPC2002).DevGUID = "{3CFA6F81-EB79-11D2-BAC5-006097BA8DF0}"
PlatType(platPC2002).ProjType = "WinCE"
PlatType(platPC2002).Platform = "{DE9660AC-85D3-4C63-A6AF-46A3B3B83737}"

cmbPlat.Clear
cmbPlat.AddItem "HPC"
cmbPlat.AddItem "Palm Size"
cmbPlat.AddItem "Pocket PC"
cmbPlat.AddItem "Pocket PC 2002"

cmbPlat.Text = "Please choose a platform ..."

End Sub
