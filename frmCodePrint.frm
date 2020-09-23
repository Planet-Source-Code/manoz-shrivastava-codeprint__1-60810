VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCodePrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Code Print"
   ClientHeight    =   5790
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   4695
   Icon            =   "frmCodePrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   4665
      TabIndex        =   3
      Top             =   5520
      Width           =   4695
   End
   Begin VB.PictureBox picTemp 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1320
      ScaleHeight     =   255
      ScaleWidth      =   1095
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.Toolbar tbr 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "iml"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "prn"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "find"
            Object.ToolTipText     =   "find"
            ImageKey        =   "find"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "exit"
            Object.ToolTipText     =   "exit"
            ImageKey        =   "close"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList iml 
         Left            =   2040
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCodePrint.frx":030A
               Key             =   "close"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCodePrint.frx":041C
               Key             =   "open"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCodePrint.frx":052E
               Key             =   "find"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCodePrint.frx":0980
               Key             =   "prn"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imlProj 
         Left            =   5040
         Top             =   720
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCodePrint.frx":0A92
               Key             =   "bas"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCodePrint.frx":0ECA
               Key             =   "dll"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCodePrint.frx":11E4
               Key             =   "ocx"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCodePrint.frx":14FE
               Key             =   "open"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCodePrint.frx":1950
               Key             =   "close"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCodePrint.frx":1DA2
               Key             =   "cls"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCodePrint.frx":21E9
               Key             =   "ctl"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCodePrint.frx":264C
               Key             =   "frm"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCodePrint.frx":2A27
               Key             =   "vbp"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.TreeView tvwVBP 
      Height          =   4935
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   8705
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin SHDocVwCtl.WebBrowser WebPreview 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   2760
      Width           =   375
      ExtentX         =   661
      ExtentY         =   661
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmCodePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ToDo list
'   - AutoFormat unformatted code
'   presently CodePrint assumes that code is already
'   formatted and proceeds with generating HTML
'   listing and displaying Print Preview screen.

Option Explicit

Public VBInstance                           As VBIDE.VBE
Public Connect                              As Connect
Private VBComp                              As VBComponent
Private VBRef                               As Reference
Private VBProj                              As VBProject
Private Const navNoHistory                  As Integer = 2
Private Const navNoWriteToCache             As Integer = 8

Private Sub Form_Load()

    On Error Resume Next
        strProjFile = VBInstance.ActiveVBProject.filename
        strProjTitle = VBInstance.ActiveVBProject.Name
        Call InitializeTree(strProjTitle)
        Call SetTreeView
        picProgress.BackColor = vbWhite
        picProgress.DrawMode = 10
        picProgress.FillStyle = 0
        picProgress.AutoRedraw = True
        lngCompCount = VBInstance.ActiveVBProject.VBComponents.Count
        SetTopMost frmCodePrint, True
        
End Sub

Private Sub GeneratePreview()

  Dim intFreeFile         As Integer
  Dim intCntr             As Integer
  Dim lngLines            As Long
  Dim cModule             As CodeModule
  Dim lngLineNo           As Long
  Dim strLine             As String
  Dim strReferences       As String
  Dim intCompType         As Integer
  Dim intQPos             As Integer
  
    On Error Resume Next
        For Each VBRef In VBInstance.ActiveVBProject.References
            strReferences = strReferences & VBRef.Name & "[" & VBRef.Description & "]" & vbNewLine
        Next
        strCode = ""
        For intCntr = 1 To tvwVBP.Nodes.Count
            DoEvents
            Call ProgressMeter(lngCompCount, CLng(intCntr), "Generating Preview (" & Round(intCntr * 100 / lngCompCount) & "%) Complete")
            If tvwVBP.Nodes(intCntr).Checked = True Then
                strFileName = tvwVBP.Nodes(intCntr).Text
                
                Set cModule = VBInstance.ActiveVBProject.VBComponents(strFileName).CodeModule
                intCompType = VBInstance.ActiveVBProject.VBComponents(strFileName).Type
                DoEvents
                strCode = strCode & vbNewLine & vbNewLine & "<HR><H4 CLASS=TITLE><CENTER>" & "[" & strFileName & " - " & ModuleTypeName(intCompType) & "]" & "</CENTER></H4><HR>" & vbNewLine
                For lngLineNo = 1 To cModule.CountOfDeclarationLines
                    If Len(cModule.Lines(lngLineNo, 1)) > 0 Then
                        strCode = strCode & vbNewLine & lngLineNo & ": " & cModule.Lines(lngLineNo, 1)
                    End If
                Next lngLineNo
                For lngLines = lngLineNo To cModule.CountOfLines
                    DoEvents
                    strCode = strCode & vbNewLine & lngLines & ": " & cModule.Lines(lngLines, 1)
                Next lngLines
            End If
            DoEvents
        Next intCntr
        intFreeFile = FreeFile
        strFileName = App.Path & "\" & VBInstance.ActiveVBProject.Name & ".html"
        Open strFileName For Output As #intFreeFile
        Print #intFreeFile, "<HTML>"
        Print #intFreeFile, "   <TITLE>Code Listing for " & VBInstance.ActiveVBProject.Name & ".vbp</TITLE>"
        Print #intFreeFile, "       <STYLE>"
        Print #intFreeFile, "           H4.TITLE {FONT WEIGHT:BOLD; BACKGROUND:WHITE; COLOR:BLACK}"
        Print #intFreeFile, "           H5.SUMMARY {FONT WEIGHT:BOLD; BACKGROUND:BLUE; COLOR:WHITE}"
        Print #intFreeFile, "       </STYLE>"
        Print #intFreeFile, "       <BODY>"
        Print #intFreeFile, "           <PRE>"
        Print #intFreeFile, "               <FONT NAME='Courier New' SIZE=1 ></FONT>"
        Print #intFreeFile, "                   <H5 CLASS=SUMMARY>References</H5>"
        Print #intFreeFile, strReferences
        Print #intFreeFile, "                   <H5 CLASS=SUMMARY>Code</H5>"
        Print #intFreeFile, strCode & vbNewLine
        Print #intFreeFile, "<H5 CLASS=SUMMARY>Code Listing Generated by " & App.ProductName & " on " & GetDate & " at " & Format(Time, "HH:MM AM/PM") & "</H5>"
        Print #intFreeFile, "               </FONT>"
        Print #intFreeFile, "           </PRE>"
        Print #intFreeFile, "       </BODY>"
        Print #intFreeFile, "</HTML>"
        Close #intFreeFile
        
        WebPreview.Navigate2 strFileName, navNoHistory & navNoWriteToCache

        picProgress.Cls
End Sub

Private Sub WebPreview_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    
    DoEvents
    On Error Resume Next
    DoEvents
End Sub

Private Sub WebPreview_DownloadComplete()
    
    On Error GoTo sub_error
    DoEvents
    If WebPreview.LocationURL <> "http:///" Then
        DoEvents
        WebPreview.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER, 0, 0
    End If
    Exit Sub
sub_error:
    DoEvents
    'ShellExecute Me.hwnd, "Open", strFileName, vbNullString, vbNullString, 1

End Sub

Private Function GetDirName() As String

    GetDirName = Mid$(strProjFile, 1, InStrRev(strProjFile, "\") - 1) & "\"

End Function

Private Sub InitializeTree(strProject As String)

    tvwVBP.Nodes.Clear
    tvwVBP.ImageList = imlProj
    Set tvProj = tvwVBP.Nodes.Add(, , "rt", strProject, ExtractIcon(strProjFile, imlProj, picTemp, 16))
    tvProj.Expanded = True
    tvProj.Bold = True
    Set tvProj = tvwVBP.Nodes.Add("rt", tvwChild, "forms", "Forms", "close", "open")
    Set tvProj = tvwVBP.Nodes.Add("rt", tvwChild, "modules", "Modules", "close", "open")
    Set tvProj = tvwVBP.Nodes.Add("rt", tvwChild, "class", "Class Modules", "close", "open")
    Set tvProj = tvwVBP.Nodes.Add("rt", tvwChild, "user", "User Control", "close", "open")
    Set tvProj = tvwVBP.Nodes.Add("rt", tvwChild, "relateddocs", "RelatedDocuments", "close", "open")

End Sub

Public Sub ProgressMeter(upperLimit As Long, progress As Long, msg As String)

  Dim r As Long

    If progress <= upperLimit Then
        If progress > picProgress.ScaleWidth Then
            progress = picProgress.ScaleWidth
        End If
        picProgress.Cls
        picProgress.ScaleWidth = upperLimit
        picProgress.CurrentX = (picProgress.ScaleWidth - picProgress.TextWidth(msg)) / 2
        picProgress.CurrentY = (picProgress.ScaleHeight - picProgress.TextHeight(msg)) / 2
        picProgress.Print msg
        picProgress.Line (0, 0)-(progress, picProgress.ScaleHeight), &HC00000, BF
        DoEvents
    End If

End Sub

Private Sub SetTreeView()

  Dim VBProj          As VBProject
  Dim VBComp          As VBComponent

    For Each VBComp In VBInstance.ActiveVBProject.VBComponents
        DoEvents
        Select Case ModuleTypeName(VBComp.Type)
          Case "Form"
            Set tvProj = tvwVBP.Nodes.Add("forms", tvwChild, , VBComp.Name, ExtractIcon(GetDirName & VBComp.Name & ".frm", imlProj, picTemp, 16))
            tvProj.Checked = True
            tvProj.Tag = GetDirName & VBComp.Name & ".frm"
          Case "Class Module"
            Set tvProj = tvwVBP.Nodes.Add("class", tvwChild, , VBComp.Name, ExtractIcon(GetDirName & VBComp.Name & ".cls", imlProj, picTemp, 16))
            tvProj.Checked = True
            tvProj.Tag = GetDirName & VBComp.Name & ".cls"
          Case "Module"
            Set tvProj = tvwVBP.Nodes.Add("modules", tvwChild, , VBComp.Name, ExtractIcon(GetDirName & VBComp.Name & ".bas", imlProj, picTemp, 16))
            tvProj.Checked = True
            tvProj.Tag = GetDirName & VBComp.Name & ".bas"
          Case "UserControl"
            Set tvProj = tvwVBP.Nodes.Add("user", tvwChild, , VBComp.Name, ExtractIcon(GetDirName & VBComp.Name & ".ctl", imlProj, picTemp, 16))
            tvProj.Checked = True
            tvProj.Tag = GetDirName & VBComp.Name & ".ctl"
          Case "RelatedDocument"
            Set tvProj = tvwVBP.Nodes.Add("relateddocs", tvwChild, , VBComp.Name, ExtractIcon(GetDirName & VBComp.Name & ".dob", imlProj, picTemp, 16))
            tvProj.Checked = True
            tvProj.Tag = GetDirName & VBComp.Name & ".dob"
          Case "Designers"
            Set tvProj = tvwVBP.Nodes.Add("relateddocs", tvwChild, , VBComp.Name, ExtractIcon(GetDirName & VBComp.Name & ".dsr", imlProj, picTemp, 16))
            tvProj.Checked = True
            tvProj.Tag = GetDirName & VBComp.Name & ".dsr"
        End Select
    Next VBComp

End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)

    If Button.Key = "exit" Then
        Connect.Hide
      ElseIf Button.Key = "print" Then
        Call GeneratePreview
        
      ElseIf Button.Key = "find" Then
    End If

End Sub

