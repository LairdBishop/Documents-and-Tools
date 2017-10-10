
Dim varOSDisk
on error resume next

Set oTSEnv = CreateObject("Microsoft.SMS.TSEnvironment")

varOSDisk=getDriveLetterFromPartNum (0, 1)

WScript.Echo  "Variable OSDisk=" & varOSDisk

'set OSDisk TS var
oTSEnv("OSDisk")=varOSDisk

WScript.Sleep (5000)
For Each oVar In oTSEnv.GetVariables
    WScript.Echo oVar & "=" & oTSEnv(oVar)
Next


Function getDriveLetterFromPartNum (iDisk, iPartNum)
	Dim query 
	Dim objWMI 
	Dim diskDrives 
	Dim diskDrive 
	Dim partitions 
	Dim partition ' will contain the drive & partition numbers
	Dim logicalDisks 
	Dim logicalDisk ' will contain the drive letter
	
	Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
	Set diskDrives = objWMI.ExecQuery("SELECT * FROM Win32_DiskDrive where Index= '" & iDisk & "'") ' First get out the physical drives
	For Each diskDrive In diskDrives 
	    query = "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" + diskDrive.DeviceID + "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition" ' link the physical drives to the partitions
	    Set partitions = objWMI.ExecQuery(query) 
	    For Each partition In partitions 
	   	WScript.Echo partition.DeviceID
	     If Trim (partition.DeviceID)= "Disk #" & iDisk & ", " & "Partition #" & iPartNum Then
	     
	        query = "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" + partition.DeviceID + "'} WHERE AssocClass = Win32_LogicalDiskToPartition"  ' link the partitions to the logical disks 
	        Set logicalDisks = objWMI.ExecQuery (query) 
	        For Each logicalDisk In logicalDisks
	        	WScript.Echo "Drive Letter associated to Disk #" & iDisk & ", " & "Partition #" & iPartNum +1 & " is " & logicalDisk.DeviceID
	            getDriveLetterFromPartNum= logicalDisk.DeviceID
	        Next
	        
	     End If
	    Next 
	Next
	
End Function