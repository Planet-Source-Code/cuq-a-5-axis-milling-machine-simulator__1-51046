Attribute VB_Name = "Systeme"

'======================================================================
'=========================  INFOS SYSTEM ==============================
'======================================================================

Public Type SYSTEM_INFO       ' Pour API
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Public SysInfo As SYSTEM_INFO

'---------- Infos System ----------

Public Sub LectureInfosSystem()
   GetSystemInfo SysInfo
End Sub
