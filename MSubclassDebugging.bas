Attribute VB_Name = "MSubclass"
' Copyright © 2017 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' MSubclassDebugging.bas
'
' Subclassing Routines
'   - Dependancies: ISubclass.cls or VB6.tlb
'   - STOP/END/PAUSE works when compiled unlike DbgWProc.dll (Curland)
'   - Multiple Subclassing of the same window
'   - Adapted from HookXP by Karl E. Peterson - http://vb.mvps.org/samples/HookXP/
'   - No SetProp/GetProp Bugs like SubTimer.dll(McKinney), SSubTmr6.dll(McMahon)
'   - No DEP issues or Assembly Thunks PushParamThunk(Curland) / SelfSub(Caton)
'   - No UserObject leaks http://stackoverflow.com/questions/1120337/setwindowsubclass-is-leaking-user-objects
'
Option Explicit

Private Enum VbDebugMode
    vbStopped
    vbRunning
    vbBreak
#If False Then
    Dim vbStopped, vbRunning, vbBreak
#End If
End Enum

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function SetWindowSubclass Lib "comctl32" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Public Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function EbMode Lib "vba6" () As Long
Private Declare Function EbIsResetting Lib "vba6" () As Long

Private Function InIDE() As Boolean
' InIDE will only return true if we have the VB's IDE DLL loaded
    Static Init As Boolean, InIDE_ As Boolean
    If Not Init Then
        Init = True
        InIDE_ = GetModuleHandle(StrPtr("vba6"))
    End If
    InIDE = InIDE_
End Function

Public Function SetSubclass(ByVal hWnd As Long, ByVal This As ISubclass, Optional ByVal dwRefData As Long) As Long
    SetSubclass = SetWindowSubclass(hWnd, AddressOf MSubclass.SubclassProc, ObjPtr(This), dwRefData)
End Function

Public Function RemoveSubclass(ByVal hWnd As Long, ByVal This As ISubclass) As Long
    RemoveSubclass = RemoveWindowSubclass(hWnd, AddressOf MSubclass.SubclassProc, ObjPtr(This))
End Function

Private Function SubclassProc(ByVal hWnd As Long, _
                              ByVal uMsg As Long, _
                              ByVal wParam As Long, _
                              ByVal lParam As Long, _
                              ByVal uIdSubclass As ISubclass, _
                              ByVal dwRefData As Long _
                              ) As Long
    
    Const WM_NCDESTROY As Long = &H82&
    
    If uMsg = WM_NCDESTROY Then
        ' We still pass WM_NCDESTROY to the subclass proc, before removing ourselves
        SubclassProc = uIdSubclass.SubclassProc(hWnd, uMsg, wParam, lParam, dwRefData)
        RemoveSubclass hWnd, uIdSubclass
        'SubclassProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
        Exit Function
    End If
    
    If Not InIDE() Then
        SubclassProc = uIdSubclass.SubclassProc(hWnd, uMsg, wParam, lParam, dwRefData)
    Else
        If EbIsResetting() Then
            RemoveSubclass hWnd, uIdSubclass
        Else
            Select Case EbMode()
            
                Case vbStopped
                    RemoveSubclass hWnd, uIdSubclass
                    
                Case vbRunning
                    ' An unhandled exception here will take down the IDE...
                    ' This happens when ebProjectReset (Stop/End) is called
                    ' during a Break (Pause) in the subclassing procedure,
                    ' or when attempting to forward to an invalid ISubclass
                    On Error Resume Next
                    SubclassProc = uIdSubclass.SubclassProc(hWnd, uMsg, wParam, lParam, dwRefData)
                    If Err Then RemoveSubclass hWnd, uIdSubclass
                    On Error GoTo 0
                    Exit Function
                    
                Case vbBreak
            
            End Select
        End If
        
        SubclassProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
    End If
End Function
