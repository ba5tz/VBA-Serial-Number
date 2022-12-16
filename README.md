[![Author](https://img.shields.io/badge/author-Andi%20Setiadi-lightgrey.svg?colorB=1D63DC&style=flat-square)]()
[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://GitHub.com/ba5tz/StrapDown.js/graphs/commit-activity)
[![Ask Me Anything !](https://img.shields.io/badge/Ask%20me-anything-1abc9c.svg)](https://s.id/setiadi)

# VBA-Serial-Number
Kumpulan Serial Number untuk kebutuhan VBA

- Serial Number Hardisk (Logical) versi FSO  
- Serial Number Hardisk (Logical) versi MSO
- Serial Number Hardisk (Fisik)
- Serial Number BIOS
- Serial Number Processor
- Serial Number MotherBoard

---
### 1. Serial Number Hardisk (Volume Serial Number | Logical)
mendapatkan Volume Serial Number hardisk melalui jalur File System Object 
```VB
Function HDSerialNumber() As String
Dim fsObj As Object, drv As Object

Set fsObj = CreateObject("Scripting.FileSystemObject")
Set drv = fsObj.Drives("C")
HDSerialNumber = drv.SerialNumber
End Function
```
#### Contoh hasil (Long Integer)
```
1552956067
```
---
### 2. Serial Number Hardisk Logical (Volume Serial Number | Logical)
mendapatkan Volume Serial Number hardisk melalui jalur Windows Management Instrumens
```VB
Public Function HDSerialNumberL() As String
Dim objWMI As Object
Dim objWin32 As Object
Dim objLD As Object
Dim strSN As Variant

Set objWMI = GetObject("WinMgmts:")
Set objWin32 = objWMI.InstancesOf("Win32_LogicalDisk")

For Each objLD In objWin32
    If objLD.DeviceID = "C:" Then
        strSN = objLD.VolumeSerialNumber
    End If
Next

HDSerialNumberL = strSN
End Function
```
#### Contoh hasil (Hex)
```
5C903AA3
```
---
### 3. Serial Number Hardisk Fisik
```VB
Public Function HDSerial() As String

Dim objWMI As Object
Dim objWin32 As Object
Dim objPM As Object
Dim strSN As String

Set objWMI = GetObject("WinMgmts:")
Set objWin32 = objWMI.InstancesOf("Win32_PhysicalMedia")

For Each objPM In objWin32
    strSN = strSN & (":" + objPM.SerialNumber)
Next

HDSerial = Trim(Mid(strSN, 2))
    
End Function
```
#### Contoh Hasil (2 Hard drive)
```
KO20210707629:AA202208201518
```

---
### 4. Serial Number BIOS
```VB
Function GetBIOSSerialNumber() As String
   
    Dim oWMI As Object    'WMI object to query
    Dim oBIOSs As Object
    Dim oBIOS As Object
    Dim Tmp As String

    Set oWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set oBIOSs = oWMI.ExecQuery("SELECT SerialNumber FROM Win32_BIOS")

    For Each oBIOS In oBIOSs
        Tmp = Tmp & ("," + oBIOS.SerialNumber)
    Next
    
    GetBIOSSerialNumber = Mid(Tmp, 2)
End Function
```
#### Contoh Hasil
```
PC027CM8
```
---
### 5. Serial Number Processor 
```vb
Function ProcessorNumber()
Dim WMI As Object
Dim Proc As Object
Dim Procs As Object
Dim Tmp As String

Set WMI = GetObject("winmgmts:")
Set Procs = WMI.ExecQuery("select ProcessorId from win32_processor")
For Each Proc In Procs
    Tmp = Tmp & ("," + Proc.ProcessorId)
Next Proc
ProcessorNumber = Mid(Tmp, 2)
End Function
```
#### Contoh hasil
```
BFEBFBFF000306C3
```

---
### 6. Serial Number MotherBoard
```vb
Function BoardSerialNumber() As String
Dim objWMI As Object
Dim objWin32 As Object
Dim Board As Object
Dim Tmp As String
        
Set objWMI = GetObject("WinMgmts:")
Set objWin32 = objWMI.InstancesOf("Win32_BaseBoard")

For Each Board In objWin32
    Tmp = Tmp & (", " + Board.SerialNumber)
Next Board
        
BoardSerialNumber = Trim(Mid(Tmp, 2))
End Function
```

#### Contoh hasil
```
L1HF4AM03A6
```
