Attribute VB_Name = "XPTest"
Option Explicit

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
        (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformID As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Public Declare Function IsAppThemed Lib "uxtheme.dll" () As Long
Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformID As Long
End Type
Private Declare Function DllGetVersion Lib "comctl32" (pdvi As DLLVERSIONINFO) As Long

Public Function IsThemedXP() As Boolean

    'Declare structure.
    Dim osVer As OSVERSIONINFO
    
    'Set size of structure.
    osVer.dwOSVersionInfoSize = Len(osVer)
    
    'Fill structure with data.
    GetVersionEx osVer
    
    'Evaluate return. If greater than or equal to 5.1 then running
    'WindowsXP or newer.
    If osVer.dwMajorVersion + osVer.dwMinorVersion / 10 >= 5.1 Then
        'Check for Active Visual Style(modified as per paravoid's suggestion).
        If IsAppThemed Then
            'Double Check by assessing DLL version loaded
            Dim dllVer As DLLVERSIONINFO
            dllVer.cbSize = Len(dllVer)
            DllGetVersion dllVer
            IsThemedXP = (dllVer.dwMajor >= 6)
        End If
    End If
    
End Function

