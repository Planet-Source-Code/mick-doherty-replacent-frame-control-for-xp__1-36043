Attribute VB_Name = "XPInit"
Option Explicit

Private Declare Function InitCommonControlsEx Lib "COMCTL32" (init As INITCC) As Boolean

Private Type INITCC
    dwSize As Long
    dwICC As Long
End Type

Private Const ICC_ANIMATE_CLASS = &H80
Private Const ICC_BAR_CLASSES = &H4
Private Const ICC_COOL_CLASSES = &H400
Private Const ICC_DATE_CLASSES = &H100
Private Const ICC_HOTKEY_CLASS = &H40
Private Const ICC_INTERNET_CLASSES = &H800
Private Const ICC_LISTVIEW_CLASSES = &H1
Private Const ICC_PAGESCROLLER_CLASS = &H1000
Private Const ICC_PROGRESS_CLASS = &H20
Private Const ICC_TAB_CLASSES = &H8
Private Const ICC_TREEVIEW_CLASSES = &H2
Private Const ICC_UPDOWN_CLASS = &H10
Private Const ICC_USEREX_CLASSES = &H200
Private Const ICC_WIN95_CLASSES = &HFF

Sub Main()
    
    Dim ICC As INITCC
    
    ICC.dwSize = Len(ICC)
    
    ICC.dwICC = ICC_ANIMATE_CLASS Or ICC_BAR_CLASSES Or ICC_COOL_CLASSES Or _
                ICC_DATE_CLASSES Or ICC_HOTKEY_CLASS Or ICC_INTERNET_CLASSES Or _
                ICC_LISTVIEW_CLASSES Or ICC_PAGESCROLLER_CLASS Or _
                ICC_PROGRESS_CLASS Or ICC_TAB_CLASSES Or ICC_TREEVIEW_CLASSES Or _
                ICC_UPDOWN_CLASS Or ICC_USEREX_CLASSES Or ICC_WIN95_CLASSES

    InitCommonControlsEx ICC
    frmMain.Show
    
End Sub
