Attribute VB_Name = "mMemory"
Option Explicit

' mdlMemory ---------------------------------
'Private Declare Function PdhOpenQuery Lib "PDH.DLL" (ByVal Reserved As Long, ByVal dwUserData As Long, ByRef hQuery As Long) As PDH_STATUS
'Private Declare Function PdhVbAddCounter Lib "PDH.DLL" (ByVal QueryHandle As Long, ByVal CounterPath As String, ByRef CounterHandle As Long) As PDH_STATUS
'Private Declare Function PdhCollectQueryData Lib "PDH.DLL" (ByVal QueryHandle As Long) As PDH_STATUS
'Private Declare Function PdhVbGetDoubleCounterValue Lib "PDH.DLL" (ByVal CounterHandle As Long, ByRef CounterStatus As Long) As Double

'CPU Information (Using: Windows Performance Data Helper DLL)
'Private Enum PDH_STATUS
'    PDH_CSTATUS_VALID_DATA = &H0
'    PDH_CSTATUS_NEW_DATA = &H1
'End Enum
'
'Private Type CounterInfo
'    hCounter As Long
'    strName As String
'End Type
'
'Dim pdhStatus As PDH_STATUS
'Dim Counters(0 To 99) As CounterInfo
'Dim hQuery As Long
'Private QueryObject As Object

'Private Type MEMORYSTATUS
'    dwLength As Long
'    dwMemoryLoad As Long
'    dwTotalPhys As Long
'    dwAvailPhys As Long
'    dwTotalPageFile As Long
'    dwAvailPageFile As Long
'    dwTotalVirtual As Long
'    dwAvailVirtual As Long
'End Type

'Private Declare Sub GlobalMemoryStatus Lib "kernel32.dll" (ByRef lpBuffer As MEMORYSTATUS)
Private Declare Function GetProcessMemoryInfo Lib "Psapi.dll" (ByVal Process As Long, ByRef ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long

Private Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type

'------------------------------------------------------------------
Private Type PerformanceStructure
    cd As Long
    commitTotal As Long
    commitLimit As Long
    commitPeak As Long
    PhysicalTotal As Long
    PhysicalAvailable As Long
    SystemCache As Long
    KernelTotal As Long
    KernelPaged As Long
    KernelNonpaged As Long
    PageSize As Long
    HandleCount As Long
    ProcessCount As Long
    ThreadCount As Long
End Type
'
'The API call to use for getting important system information related to memory.
Private Declare Function GetPerformanceInfo Lib "Psapi.dll" (performanceInfo As PerformanceStructure, ByVal structureSize As Long) As Boolean

Public Function GetMemory(ProcessID As Long) As String
    On Error Resume Next
    Dim byteSize As Double, hProcess As Long, ProcMem As PROCESS_MEMORY_COUNTERS
    ProcMem.cb = LenB(ProcMem)
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessID)
    If hProcess <= 0 Then GetMemory = "N/A": Exit Function
    GetProcessMemoryInfo hProcess, ProcMem, ProcMem.cb
    byteSize = ProcMem.WorkingSetSize
    GetMemory = byteSize
    Call CloseHandle(hProcess)
End Function

Sub MonitoringPerformance(lblHandles As Label, lblTotalPhysMem As Label, lblAvaiMem As Label, _
                            lblSystemCache As Label, lblTotalCommit As Label, lblLimit As Label, _
                            lblPeak As Label, lblTotalKernelMem As Label, lblPaged As Label, _
                            lblNonPaged As Label, lblThread As Label, lblProcess As Label)
    
    'Declare a variable to access the structure which will be passed to the API call.
    Dim ps As PerformanceStructure
    'The return value from the API call. True means it should be OK, false means otherwise.
    Dim ret As Boolean
    'Call the API with the ps structure variable.
    ret = GetPerformanceInfo(ps, Len(ps))
    'Page size variable which is required to accurately calculate the memory
    Dim page As Long
    'Get the PageSize from the structure and calculate to KBytes
    page = ps.PageSize / 1024
    
    'The rest of code is retrieving the info from the structure which was passed to the API call.
'    DoEvents
    lblHandles.Caption = ps.HandleCount
    lblTotalPhysMem.Caption = ps.PhysicalTotal * page
    lblAvaiMem.Caption = ps.PhysicalAvailable * page
    lblSystemCache.Caption = ps.SystemCache * page
    lblTotalCommit.Caption = ps.commitTotal * page
    lblLimit.Caption = ps.commitLimit * page
    lblPeak.Caption = ps.commitPeak * page
    lblTotalKernelMem.Caption = ps.KernelTotal * page
    lblPaged.Caption = ps.KernelPaged * page
    lblNonPaged.Caption = ps.KernelNonpaged * page
    lblThread.Caption = ps.ThreadCount
    lblProcess.Caption = ps.ProcessCount
'    "PF Usage : " & Format(ps.commitTotal * page \ 1024, "") & " MB"
End Sub

'Sub MemoryInfo(lbPhysMem As Label, lbAvaiPhyMem As Label, lbUsedPhyMem As Label, _
'                    lbMemLoad As Label, lbPagFile As Label, lbAvaiPagFile As Label, _
'                    lbPagFileUsg As Label, lbAvaiVirMem As Label, lbUsedVirMem As Label, _
'                    lbVirMem As Label, lbTotal As Label)
'
'    Dim mem As MEMORYSTATUS
'    mem.dwLength = Len(mem)
'    GlobalMemoryStatus mem
'
'    lbPhysMem.Caption = Format(mem.dwTotalPhys \ 1024, "")
'    lbAvaiPhyMem.Caption = Format(mem.dwAvailPhys \ 1024, "")
'    lbUsedPhyMem.Caption = Format(lbAvaiPhyMem.Caption / lbPhysMem.Caption * 100, "0.00") & " %"
'    lbPagFile.Caption = Format(mem.dwTotalPageFile \ 1024, "")
'    lbAvaiPagFile.Caption = Format(mem.dwAvailPageFile \ 1024, "")
'    lbPagFileUsg.Caption = Format(lbAvaiPagFile.Caption / lbPagFile.Caption * 100, "0.00") & " %"
'    lbVirMem.Caption = Format(mem.dwTotalVirtual \ 1024, "")
'    lbAvaiVirMem.Caption = Format(mem.dwAvailVirtual \ 1024, "")
'    lbUsedVirMem.Caption = Format(lbAvaiVirMem.Caption / lbVirMem.Caption * 100, "0.00") & " %"
'    lbTotal.Caption = Format(mem.dwTotalPageFile \ 1024 - mem.dwAvailPageFile \ 1024, "")
'End Sub

'New for monitoring performance windows like Task Manager Windows
'Update 15 Maret 2009 16:30

'Sub GetCPUInfo(sbCPU As StatusBar)
'    pdhStatus = PdhOpenQuery(0, 1, hQuery)
'    AddCounter "\Processor(0)\% Processor Time", hQuery
'    UpdateValues sbCPU
'End Sub
'
'Sub UpdateValues(sbCPU As StatusBar)
'    Dim dblCounterValue As Double
'    Dim pdhStatus As Long
'    Dim strInfo As String
'    Dim i As Long
'    PdhCollectQueryData (hQuery)
'    i = 0
'    dblCounterValue = PdhVbGetDoubleCounterValue(Counters(i).hCounter, pdhStatus)
'    If (pdhStatus = PDH_CSTATUS_VALID_DATA) Or (pdhStatus = PDH_CSTATUS_NEW_DATA) Then
'        sbCPU.Panels(2).Text = "Processor Usage : " & Format(dblCounterValue, "0") & " %"
'    End If
'End Sub
'
'Sub AddCounter(strCounterName As String, hQuery As Long)
'    Dim pdhStatus As PDH_STATUS
'    Dim hCounter As Long, currentCounterIdx As Long
'
'    pdhStatus = PdhVbAddCounter(hQuery, strCounterName, hCounter)
'    Counters(currentCounterIdx).hCounter = hCounter
'    Counters(currentCounterIdx).strName = strCounterName
'    currentCounterIdx = currentCounterIdx + 1
'End Sub
