VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   340
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   496
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.TreeView tvw 
      Height          =   4215
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   7435
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   255
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":09F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0D48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":109A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":13EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":173E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1DE2
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Scopes
Private Const RESOURCE_CONNECTED = &H1
Private Const RESOURCE_ENUM_ALL = &HFFFF
Private Const RESOURCE_GLOBALNET = &H2
Private Const RESOURCE_REMEMBERED = &H3

'The order of images in image list are based off these values
Private Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Private Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Private Const RESOURCEDISPLAYTYPE_SERVER = &H2
Private Const RESOURCEDISPLAYTYPE_SHARE = &H3
Private Const RESOURCEDISPLAYTYPE_FILE = &H4
Private Const RESOURCEDISPLAYTYPE_GROUP = &H5
Private Const RESOURCEDISPLAYTYPE_NETWORK = &H6
Private Const RESOURCEDISPLAYTYPE_ROOT = &H7
Private Const RESOURCEDISPLAYTYPE_SHAREADMIN = &H8
Private Const RESOURCEDISPLAYTYPE_DIRECTORY = &H9
Private Const RESOURCEDISPLAYTYPE_TREE = &HA
Private Const RESOURCEDISPLAYTYPE_NDSCONTAINER = &HB


Private Const RESOURCETYPE_ANY = &H0
Private Const RESOURCETYPE_DISK = &H1
Private Const RESOURCETYPE_PRINT = &H2
Private Const RESOURCETYPE_UNKNOWN = &HFFFF

Private Const RESOURCEUSAGE_ALL = &H0
Private Const RESOURCEUSAGE_CONNECTABLE = &H1
Private Const RESOURCEUSAGE_CONTAINER = &H2
Private Const RESOURCEUSAGE_RESERVED = &H80000000

Private Const ERROR_NO_MORE_ITEMS = 259

Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As Long
    lpRemoteName As Long
    lpComment As Long
    lpProvider As Long
End Type
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, lpBuffer As Any, lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long

Private Sub Form_Load()
    
    OpenNetworkResources

End Sub

Private Sub OpenNetworkResources()

    Dim lngresult As Long
    Dim lngenumhwnd As Long
    Dim lngentries As Long
    Dim i As Integer
    Dim strremotename As String
    
    Dim netdata(511) As NETRESOURCE
    lngentries = -1
    
    tvw.Nodes.Clear
    
    lngresult = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, RESOURCEUSAGE_ALL, ByVal 0, lngenumhwnd)
    If lngresult = 0 And lngenumhwnd <> 0 Then
        lngresult = WNetEnumResource(lngenumhwnd, lngentries, netdata(0), CLng(Len(netdata(0))) * 512)
        If lngresult = 0 Then
            For i = 0 To lngentries - 1
                strremotename = ltos(netdata(i).lpRemoteName)
                strremotename = ParseName(strremotename)
                'These are separated for different images
                tvw.Nodes.Add(, , strremotename, strremotename, netdata(i).dwDisplayType + 1).Expanded = True
                If netdata(i).dwUsage And RESOURCEUSAGE_CONTAINER Then
                    EnumerateNetworkResources netdata(i), strremotename 'enumerate subitems
                End If
            Next i
        ElseIf lngresult = ERROR_NO_MORE_ITEMS Then
        Else
            'MsgBox "ERROOOOOOOORRRRRGH! " + CStr(lngresult)
        End If
    Else
        'MsgBox "ERROOOOOOOORRRRRGH! " + CStr(lngresult)
    End If
    lngresult = WNetCloseEnum(lngenumhwnd)


End Sub

Private Sub EnumerateNetworkResources(netdata_parent As NETRESOURCE, strremotename_parent As String)

    Dim lngresult As Long
    Dim lngenumhwnd As Long
    Dim lngentries As Long
    Dim i As Integer
    Dim strremotename As String
    
    Dim netdata(511) As NETRESOURCE
    lngentries = -1
    
    lngresult = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, RESOURCEUSAGE_ALL, netdata_parent, lngenumhwnd)
    If lngresult = 0 And lngenumhwnd <> 0 Then
        lngresult = WNetEnumResource(lngenumhwnd, lngentries, netdata(0), CLng(Len(netdata(0))) * 512)
        If lngresult = 0 Then
            For i = 0 To lngentries - 1
                strremotename = ltos(netdata(i).lpRemoteName)
                strremotename = ParseName(strremotename)
                'These are separated for different images
                tvw.Nodes.Add(strremotename_parent, 4, strremotename_parent + strremotename, strremotename, netdata(i).dwDisplayType + 1).Expanded = True
                If netdata(i).dwUsage And RESOURCEUSAGE_CONTAINER Then
                    EnumerateNetworkResources netdata(i), strremotename_parent + strremotename 'enumerate more subitems
                End If
            Next i
        ElseIf lngresult = ERROR_NO_MORE_ITEMS Then
        Else
            'MsgBox "ERROOOOOOOORRRRRGH! " + CStr(lngresult)
        End If
    Else
        'MsgBox "ERROOOOOOOORRRRRGH! " + CStr(lngresult)
    End If
    lngresult = WNetCloseEnum(lngenumhwnd)

End Sub

Private Function ltos(lngh As Long) As String

    Dim strl As String
    strl = Space(lstrlen(lngh))
    lstrcpy strl, lngh
    ltos = strl

End Function

Public Function ParseName(strpath As String) As String

    On Local Error Resume Next
    Dim intseppos As Integer
    intseppos = InStrRev(strpath, "\")
        ParseName = strpath
    If intseppos > 0 Then
        ParseName = Right(strpath, Len(strpath) - intseppos)
    End If

End Function
