Attribute VB_Name = "Module1"
'Option Explicit

Public Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 0) As SAFEARRAYBOUND
End Type

Public Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type


Public Declare Function GetObjectAPI _
Lib "gdi32" Alias "GetObjectA" _
(ByVal hObject As Long, _
ByVal nCount As Long, lpObject As Any) As Long

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Declare Function VarPtrArray Lib _
"msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

