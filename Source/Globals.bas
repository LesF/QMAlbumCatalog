Attribute VB_Name = "Globals"
Option Explicit

'----- API Service parameters
Public SVC_Host As String
Public SVC_Port As Long
Public SVC_User As String
Public SVC_Pass As String
Public SVC_Account As String

'----- ListView management stuff
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)

Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, _
     ByVal wParam As Long, ByVal lParam As Long) As Long

Public Sub AutoSizeListViewColumns(ByRef lvw As ListView, _
                                   Optional ByVal UseHeader As Boolean = True)
    '-----
    Dim i As Long
    Dim mode As Long
    '-----
    If UseHeader Then
        mode = LVSCW_AUTOSIZE_USEHEADER
    Else
        mode = LVSCW_AUTOSIZE
    End If

    For i = 0 To lvw.ColumnHeaders.Count - 1
        Call SendMessage(lvw.hWnd, LVM_SETCOLUMNWIDTH, i, mode)
    Next i
End Sub

Public Sub AutoSizeListViewColumnsFillLast(ByRef lvw As ListView, _
                                           Optional ByVal Padding As Long = 12)
    '-----
    Dim i As Long
    Dim totalWidth As Long
    Dim lastWidth As Long
    '-----
    ' First auto-size all columns normally
    AutoSizeListViewColumns lvw

    ' Sum widths of all columns except the last
    For i = 1 To lvw.ColumnHeaders.Count - 1
        totalWidth = totalWidth + lvw.ColumnHeaders(i).Width
    Next i

    ' Calculate remaining space inside the ListView
    lastWidth = lvw.Width - totalWidth - Padding

    ' Prevent negative or tiny widths
    If lastWidth < 100 Then lastWidth = 100

    ' Apply width to last column
    lvw.ColumnHeaders(lvw.ColumnHeaders.Count).Width = lastWidth
End Sub


