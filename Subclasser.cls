VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Subclasser"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Copyright © 2017 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' Subclasser.cls
'
' Global Subclassing Interface
'   - Dependancy: MSubclassDebugging.bas, ISubclass.cls or VB6.tlb
'
Option Explicit

Public Function DefSubclassProc(ByVal hWnd As Long, _
                                ByVal uMsg As Long, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long _
                                ) As Long
    DefSubclassProc = MSubclass.DefSubclassProc(hWnd, uMsg, wParam, lParam)
End Function

Public Function SetSubclass(ByVal hWnd As Long, ByVal This As ISubclass, Optional ByVal dwRefData As Long) As Long
    SetSubclass = MSubclass.SetSubclass(hWnd, This, dwRefData)
End Function

Public Function RemoveSubclass(ByVal hWnd As Long, ByVal This As ISubclass) As Long
    RemoveSubclass = MSubclass.RemoveSubclass(hWnd, This)
End Function

