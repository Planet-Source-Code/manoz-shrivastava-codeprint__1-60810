Attribute VB_Name = "modCodePrint"
Option Explicit

Public lngCompCount                         As Long
Public strProjFile                          As String
Public strProjTitle                         As String
Public strProjPath                          As String
Public tvProj                               As Node
Public strFileName                          As String
Public intCount                             As Long
Public strCode                              As String
Public Const LVM_FIRST                      As Long = &H1000
Public Const LVM_SETCOLUMNWIDTH             As Long = (LVM_FIRST + 30)
Public Const LVSCW_AUTOSIZE                 As Long = -1
Public Const LVSCW_AUTOSIZE_USEHEADER       As Long = -2
Private Const EM_LINESCROLL = &HB6
Private Const SWP_NOACTIVATE        As Long = &H10
Private Const SWP_NOSIZE            As Long = &H1
Private Const SWP_NOMOVE            As Long = &H2
Private Const HWND_TOPMOST          As Long = (-1)
Private Const HWND_NOTOPMOST        As Long = (-2)

Private Type typSHFILEINFO
    hIcon                                     As Long
    iIcon                                     As Long
    dwAttributes                              As Long
    szDisplayName                             As String * 260
    szTypeName                                As String * 80
End Type

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Private FileInfo                            As typSHFILEINFO

Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400
Const CB_FINDSTRING = &H14C
Const CB_FINDSTRINGEXACT = &H158
Const CB_SHOWDROPDOWN = &H14F

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, _
                                                                                 ByVal dwFileAttributes As Long, _
                                                                                 psfi As typSHFILEINFO, _
                                                                                 ByVal cbSizeFileInfo As Long, _
                                                                                 ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, _
                                                            ByVal I&, _
                                                            ByVal hDCDest&, _
                                                            ByVal x&, _
                                                            ByVal y&, _
                                                            ByVal Flags&) As Long


Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                   ByVal hWndInsertAfter As Long, _
                                                   ByVal x As Long, _
                                                   ByVal y As Long, _
                                                   ByVal cx As Long, _
                                                   ByVal cy As Long, _
                                                   ByVal wFlags As Long) As Long

Public Function ExtractIcon(filename As String, _
                            AddtoImageList As ImageList, _
                            PictureBox As PictureBox, _
                            PixelsXY As Integer) As Long

Dim SmallIcon As Long
Dim NewImage  As ListImage
Dim IconIndex As Integer

    If PixelsXY = 16 Then
        SmallIcon = SHGetFileInfo(filename, 0&, _
                    FileInfo, Len(FileInfo), Flags Or SHGFI_SMALLICON)
    Else
        SmallIcon = SHGetFileInfo(filename, 0&, _
                    FileInfo, Len(FileInfo), Flags Or SHGFI_LARGEICON)
    End If
    If SmallIcon <> 0 Then
        With PictureBox
            .Height = 15 * PixelsXY
            .Width = 15 * PixelsXY
            .ScaleHeight = 15 * PixelsXY
            .ScaleWidth = 15 * PixelsXY
            .Picture = LoadPicture("")
            .AutoRedraw = True
            SmallIcon = ImageList_Draw(SmallIcon, FileInfo.iIcon, .hDC, 0, 0, ILD_TRANSPARENT)
            .Refresh
        End With
        IconIndex = AddtoImageList.ListImages.Count + 1
        Set NewImage = AddtoImageList.ListImages.Add(IconIndex, , PictureBox.Image)
        ExtractIcon = IconIndex
    End If

End Function

Public Function IsFileExist(strFileName As String) As Boolean

    On Error GoTo err_handler
    Call FileLen(strFileName)
    IsFileExist = True

Exit Function

err_handler:
    IsFileExist = False

End Function

'Public Function LvAutoSize(lv As ListView)
'
'Dim Col2Adjst As Long
'Dim LngRc     As Long
'
'    For Col2Adjst = 0 To lv.ColumnHeaders.Count - 1
'        LngRc = SendMessage(lv.hwnd, LVM_SETCOLUMNWIDTH, Col2Adjst, ByVal LVSCW_AUTOSIZE_USEHEADER)
'    Next Col2Adjst
'
'End Function

Public Function ModuleTypeName(cmpType As Integer) As String

    On Error Resume Next
        DoEvents
        Select Case cmpType
        Case vbext_ct_StdModule
            ModuleTypeName = "Module"
        Case vbext_ct_ClassModule
            ModuleTypeName = "Class Module"
        Case vbext_ct_MSForm
            ModuleTypeName = "Form"
        Case vbext_ct_ResFile
            ModuleTypeName = "Resource"
        Case vbext_ct_VBForm
            ModuleTypeName = "Form"
        Case vbext_ct_VBMDIForm
            ModuleTypeName = "MDIForm"
        Case vbext_ct_PropPage
            ModuleTypeName = "Property Page"
        Case vbext_ct_UserControl
            ModuleTypeName = "UserControl"
        Case vbext_ct_DocObject
            ModuleTypeName = "RelatedDocument"
        Case vbext_ct_ActiveXDesigner
            ModuleTypeName = "Designers"
        End Select

End Function

Public Sub SetTopMost(frm1 As Form, _
                      ByVal isTopMost As Boolean)

    SetWindowPos frm1.hwnd, IIf(isTopMost, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
    DoEvents

End Sub

Public Function GetDate() As String
    Dim strDate As String
    Dim intDate As Integer
    
    intDate = Format(Date, "DD")
    Select Case intDate
        Case 1, 21, 31
            strDate = intDate & "st"
        Case 2, 22
            strDate = intDate & "nd"
        Case 3, 23
            strDate = intDate & "rd"
        Case Else
            strDate = intDate & "th"
    End Select
    
    strDate = strDate & " " & Format(Date, "MMM, YYYY")
    GetDate = strDate
End Function
