<#
The Big Pipeline Solution
Version 1.0
Dec 2016

Pipeline Functions
    Get-MyServerCollection
    Get-ServerObjectCollection
    Test-ServerConnectionOnPipeline
    Get-OSCaptionOnPipeline
    Get-TimeZoneOnPipeline
    Get-TotalMemoryOnPipeline
    Get-MachineModelOnPipeline
    Get-ProcInfoOnPipeline
    Get-VolumeInfoOnPipeline

Other Functions (Called by Test-ServerConnectionOnPipeline)
    Get-WMI_OS
    Get-PSRemoteComputerName

This started out as a thought experiment in extreme pipelining. The idea is to start with
a generic collection of objects (PSObject). One column of this 'table' will be ComputerName.

The original conception (and still a very good one) was to use a CSV via Import-Csv. This file
filtered by any of the other columns in the CSV. For instance Location -eq Arizona or some such.
Get-MyServerCollection is an example of how this could work.

A late additon function, Get-ServerObjectCollection, will create a similar object as Import-Csv but
from a simple collection/array of strings that are computer names. Using Get-Content, this can 
even be a txt file of server names, one name per line.

Each subsequent function will add one of more columns to the object being passed down the pipe. At
least one function, Get-VolumeInfoOnPipeline, will add rows as well.

The Test-ServerConnectionOnPipeline will almost always be included just after one or more functions/
cmdlets gathers a collection with ComputerName. The Test function will run a series of server health
tests and will return true/false on a column for that test. Subsequent functions are then coded to 
only proceed on an action against a server if the test returns true. The current tests are Ping,
WMI and PSRemote.

The rest of the functions in this example will gather Server properties describing its capacity
such as memory, processors and disk. This solution is designed to answer questions like: Are there
any SQL servers with drives less that 5% free space? Do all of the Web Servers in Arizona have more
than 48 GB of memory? You are handed a seemingly random list of servers and asked what Make and Model
of hardware are they?

The output of these functions can be piped through the Cmdlets Select-Object, Where-Object, Sort-Object
or sent to Export-Csv to be saved or later opened in Excel
ot sent to the Out-GridView to be tweaked immediately and interactively.

Future ideas/improvements

1. A bulk server update with an 'undo button'
    Start with a Registry change. Store the path, item, original value and success/fail for each change
    Store this in a CSV file. A function could be built to set all values in the CSV file back to its
    original value.
#>


<# 
.Synopsis 
   Gets a (filtered) list of servers from a CSV File 
.DESCRIPTION 
   The parameters are both optional. 
   Leaving one blank applies no filter for that parameter. 
.EXAMPLE 
   Get-MyServerCollection 
   Returns everything 
.EXAMPLE 
   Get-MyServerCollection -Role Web 
   Returns all of the Web Servers 
.EXAMPLE 
   Get-MyServerCollection -Role SQL -Location WA 
   Returns the SQL Servers in Washington 
#>
Function Get-MyServerCollection  
    { 
    Param 
        ( 
        [ValidateSet("Web", "SQL", "DC")] 
        [string]$Role, 
         
        [ValidateSet("AZ", "WA")] 
        [string]$Location 
        ) 

    # $ScriptPath = 'C:\Scripts\Book\Chap2' 
    $ScriptPath = $PSScriptRoot
    $ComputerNames = 'Servers.csv' 

    If ($Role -ne "")  {$ModRole = $Role} 
        Else {$ModRole = "*"} 
    If ($Location -ne "")  {$ModLocation = $Location} 
        Else {$ModLocation = "*"} 

    Import-Csv -Path "$ScriptPath\$ComputerNames"  | 
        Where {($_.Role -like $ModRole) -and ($_.Location -like $ModLocation)} 
    }

<# 
.Synopsis 
    Converts a collection of Server Name Strings into a Colection of objects
.DESCRIPTION 
    This function will take a random list of servers, such as an array or txt file
    and convert it on the pipeline to a collection of PSObjects. This collection 
    will function exactly like an imported CSV with ComputerName as the column heading.
.EXAMPLE 
    ('Server2','Server4') | Get-ServerObjectCollection | Test-ServerConnectionOnPipeline | ft
.EXAMPLE 
    Get-Content -Path .\RndListOfServers.txt | Get-ServerObjectCollection | Test-ServerConnectionOnPipeline | ft
.EXAMPLE 
    (Get-ADComputer -Filter *).Name | Get-ServerObjectCollection | Test-ServerConnectionOnPipeline | ft
    Active Directory!! (All Computers)
.EXAMPLE
    (Get-ADComputer -SearchBase "OU=Domain Controllers,DC=coatelab,DC=com" -Filter *).Name | Get-ServerObjectCollection | Test-ServerConnectionOnPipeline | ft
    Active Directory!! (Just Domain Controllers)
#>
Function Get-ServerObjectCollection
    {
    [CmdletBinding()]
    Param(
        [parameter(
        Mandatory=$true,
        ValueFromPipeline= $true)]
        [string]
        $ComputerName
    )

    Begin
        {}
    Process
        {
        New-Object PSObject -Property @{'ComputerName' = $_}
        }
    }

<#
Simple Function to test WMI connectivity on a remote machine 
moving the Try...Catch block into isolation helps prevent any errors on the console
Return is the WMI OS object when sucessfully connects, Null when it does not
#>
Function Get-WMI_OS ($ComputerName)
    {
    Try {Get-Wmiobject -ComputerName $ComputerName -Class Win32_OperatingSystem -ErrorAction Stop}
    Catch {}
    }

<#
Simple Function to test PS Remote connectivity on a remote machine 
moving the Try...Catch block into isolation helps prevent any errors on the console
Return is the the remote computer's name when sucessfully connects, Null when it does not
#>
Function Get-PSRemoteComputerName  ($ComputerName)
    {
    Try {Invoke-Command -ComputerName $ComputerName -ScriptBlock {1} -ErrorAction Stop}
    Catch {} 
    }

<# 
.Synopsis 
    Runs availability checks on servers
.DESCRIPTION 
    This typically takes an imported csv file with a ComputerName Column as an imput object
    but just about any collection of objects that exposes .ComputerName should work
    The output is the same type of object as the input (hopefully) so that it can be piped 
    to the next function to add another column.

    Makes both a WMI and PS Remote call
.EXAMPLE 
    Get-MyServerCollection | Test-ServerConnectionOnPipeline | ft
.EXAMPLE
    $a = Foreach ($s in ('Server1','Server2','Server3')) {New-Object PSObject -Property @{'ComputerName' = $s}}
    $a | Test-ServerConnectionOnPipeline | ft
    Another Ad Hoc way to build an object for the pipeline. These two lines cannot
#>
Function Test-ServerConnectionOnPipeline
    {
    [CmdletBinding()]

    Param
        (
        [parameter(
        Mandatory=$true, 
        ValueFromPipeline= $true)]
        $ComputerProperties
        )
    
    Begin
        {}
    Process
        {
        $ComputerProperties | Select *, Ping, WMI, PSRemote, BootTime | %{
            # Test Ping
            $_.Ping = Test-Connection -ComputerName $ComputerProperties.ComputerName -Quiet -Count 1

            If ($_.Ping)
                {
                
                # Calling WMI in a wrapper in order to isolate the error condition if it occurs
                $os = Get-WMI_OS -ComputerName $ComputerProperties.ComputerName
                # $os = Get-Wmiobject -ComputerName $ComputerProperties.ComputerName -Class Win32_OperatingSystem -ErrorAction Stop

                If ($os -ne $null) 
                    {
                    $_.BootTime = [Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime) 
                    $_.WMI = $true
                    }
                Else
                    {
                    $_.WMI = $false
                    $_.BootTime = 'No Try'
                    }

                # Test PS Remoting
                $ps = Get-PSRemoteComputerName -ComputerName $ComputerProperties.ComputerName
                # $Result = Invoke-Command -ComputerName $ComputerProperties.ComputerName -ScriptBlock {$env:COMPUTERNAME} -ErrorAction Stop

                If ($ps -ne $null) 
                    {$_.PSRemote = $true}
                Else
                    {$_.PSRemote = $false}
                }
            $_
            }
        }
    }

<# 
.Synopsis 
    Adds an OSVersion Column to an object
.DESCRIPTION 
    This typically takes an imported csv file with a ComputerName Column as an imput object
    but just about any collection of objects that exposes .ComputerName should work
    The output is the same type of object as the input (hopefully) so that it can be piped 
    to the next function to add another column

    Makes a call to WMI

    Requires Input object with Boolean Ping and WMI properties. Will only try to get value 
    if both are true
.EXAMPLE 
    Get-MyServerCollection | Test-ServerConnectionOnPipeline | Get-OSCaptionOnPipeline | Out-GridView
.EXAMPLE
    Get-MyServerCollection | Test-ServerConnectionOnPipeline | Get-OSCaptionOnPipeline | Select ComputerName, OSVersion | ft -AutoSize
    For a more concise output
#>
Function Get-OSCaptionOnPipeline
    {
    [CmdletBinding()]

    Param
        (
        [parameter(
        Mandatory=$true, 
        ValueFromPipeline= $true)]
        $ComputerProperties,

        [switch]
        $NoErrorCheck
        )
    
    Begin {}
        
    Process
        {
        $ComputerProperties | Select *, OSVersion | %{
            If ((($_.Ping) -and ($_.WMI)) -or ($NoErrorCheck))
                {
                $arr = $str = $_.OSVersion = $Version = $r = $null
                
                $str = (Get-WmiObject -class Win32_OperatingSystem -computerName $_.ComputerName).Caption
                $arr = $str.Split(' ',[System.StringSplitOptions]::RemoveEmptyEntries)
                
                Foreach ($a in $arr)
                    {
                    If ($a -eq 'R2'){$r = 'R2'}
                
                    Try
                        {$Version = [int]$a}
                    Catch{}
                    }
                
                $_.OSVersion = "$Version $r"
                }
            Else{$_.OSVersion = 'No Try'}
            $_
            }
        }
    }


<# 
.Synopsis 
     Adds a TimeZone Column to an object
.DESCRIPTION 
    This typically takes an imported csv file with a ComputerName Column as an imput object
    but just about any collection of objects that exposes .ComputerName should work
    The output is the same type of object as the input (hopefully) so that it can be piped 
    to the next function to add another column
    
    Uses PowerShell Remoting

    Requires Input object with Boolean Ping and PSRemote properties. Will only try to get value 
    if both are true
.EXAMPLE 
    Get-MyServerCollection | Test-ServerConnectionOnPipeline | Get-TimeZoneOnPipeline | ft
.EXAMPLE 
    Get-MyServerCollection | Test-ServerConnectionOnPipeline | Get-TimeZoneOnPipeline | Select ComputerName,TimeZone | ft -AutoSize
#>
Function Get-TimeZoneOnPipeline
    {
    [CmdletBinding()]

    Param
        (
        [parameter(
        Mandatory=$true, 
        ValueFromPipeline= $true)]
        $ComputerProperties,

        [switch]
        $NoErrorCheck
        )
    
    Begin
        {}
    Process
        {
        $ComputerProperties | Select *, TimeZone | %{
            If ((($_.Ping) -and ($_.PSRemote)) -or ($NoErrorCheck))
                {
                $_.TimeZone = $null
                
                $sb = {(Get-ItemProperty -Path 'HKLM:\system\CurrentControlSet\control\TimeZoneInformation'`
                     -Name TimeZoneKeyName).TimeZoneKeyName}
                $_.TimeZone = Invoke-Command -ComputerName $_.ComputerName -ScriptBlock $sb
                }
            Else {$_.TimeZone = 'No Try'}
            $_
            }
        }
    }

<# 
.Synopsis 
    Adds a TotalMemory Column to an object
.DESCRIPTION 
    This typically takes an imported csv file with a ComputerName Column as an imput object
    but just about any collection of objects that exposes .ComputerName should work
    The output is the same type of object as the input (hopefully) so that it can be piped 
    to the next function to add another column
.EXAMPLE 
    Get-MyServerCollection | Test-ServerConnectionOnPipeline | Get-TimeZoneOnPipeline | Get-TotalMemoryOnPipeline | ft
#>
Function Get-TotalMemoryOnPipeline
    {
    [CmdletBinding()]

    Param
        (
        [parameter(
        Mandatory=$true,
        ValueFromPipeline= $true)]
        $ComputerProperties,

        [switch]
        $NoErrorCheck
        )
    
    Begin
        {}
    Process
        {
        $ComputerProperties | Select *, TotalMemory | %{
            If ((($_.Ping) -and ($_.WMI)) -or ($NoErrorCheck))
                {
                $_.TotalMemory = [string](Get-WMIObject -class Win32_PhysicalMemory -ComputerName $_.ComputerName |
                    Measure-Object -Property capacity -Sum | 
                    % {[Math]::Round(($_.sum / 1GB),2)}) + ' GB'
                }
            Else{$_.TotalMemory = 'No Try'}
            $_
            }
        }
    }

<# 
.Synopsis 
    Adds a MachineModel Column to an object
.DESCRIPTION 
    This typically takes an imported csv file with a ComputerName Column as an imput object
    but just about any collection of objects that exposes .ComputerName should work
    The output is the same type of object as the input (hopefully) so that it can be piped 
    to the next function to add another column
.EXAMPLE 
    Get-MyServerCollection | Test-ServerConnectionOnPipeline | Get-TimeZoneOnPipeline | Get-TotalMemoryOnPipeline | Get-MachineModelOnPipeline | ft
#>
Function Get-MachineModelOnPipeline
    {
    [CmdletBinding()]

    Param
        (
        [parameter(
        Mandatory=$true,
        ValueFromPipeline= $true)]
        $ComputerProperties,

        [switch]
        $NoErrorCheck
        )
    
    Begin
        {}
    Process
        {
        $ComputerProperties | Select *, MachineModel | %{
            If ((($_.Ping) -and ($_.WMI)) -or ($NoErrorCheck))
                {
                $_.MachineModel = [string](Get-WMIObject -class Win32_ComputerSystem -ComputerName $_.ComputerName).Model
                }
            Else{$_.MachineModel = 'No Try'}
            $_
            }
        }
    }

<# 
.Synopsis 
    Adds ProcInfo (Physical) Columns to an object
.DESCRIPTION 
    This typically takes an imported csv file with a ComputerName Column as an imput object
    but just about any collection of objects that exposes .ComputerName should work
    The output is the same type of object as the input (hopefully) so that it can be piped 
    to the next function to add another column
    Columns returned are a subset of Win32_Processor
.EXAMPLE 
    Get-MyServerCollection | Test-ServerConnectionOnPipeline | Get-OSCaptionOnPipeline | Get-TimeZoneOnPipeline | Get-TotalMemoryOnPipeline | Get-MachineModelOnPipeline | Get-ProcInfoOnPipeline | Select ComputerName,BootTime,OSVersion,TimeZone,TotalMemory,MachineModel,TotalProcs,ProcName,Cores,DataWidth | ft
#>
Function Get-ProcInfoOnPipeline
    {
    [CmdletBinding()]

    Param
        (
        [parameter(
        Mandatory=$true,
        ValueFromPipeline= $true)]
        $ComputerProperties,

        [switch]
        $NoErrorCheck
        )
    
    Begin
        {}
    Process
        {
        $ComputerProperties | Select *, TotalProcs, ProcName, Cores, DataWidth | %{
            If ((($_.Ping) -and ($_.WMI)) -or ($NoErrorCheck))   
                {
                $Proc = (Get-WmiObject -computername $_.ComputerName -class win32_Processor)
                $_.TotalProcs = $Proc.count 

                # When there is only one Proc use the object as normal
                # Else use the first instance of the object in the collection -> '[0]'
                If ($_.TotalProcs -eq $null)
                    {
                    $_.TotalProcs = 1
                    $ProcAdjusted = ($Proc)
                    }
                Else 
                    {
                    $ProcAdjusted = ($Proc)[0]
                    }

                $_.ProcName = $ProcAdjusted.Name
                $_.Cores = $ProcAdjusted.NumberOfCores
                $_.DataWidth = $ProcAdjusted.DataWidth                
                }

           Else
                {
                $_.TotalProcs = 'No Try'
                $_.ProcName = 'No Try'
                $_.Cores = 'No Try'
                $_.DataWidth = 'No Try'
                }
            $_
            }
        }
    }

<# 
.Synopsis 
    Adds Volume Info Columns to an object
.DESCRIPTION 
    This typically takes an imported csv file with a ComputerName Column as an imput object
    but just about any collection of objects that exposes .ComputerName should work
    The output is the same type of object as the input (hopefully) so that it can be piped 
    to the next function to add another column
    Columns returned are a subset of Win32_Volume

    Adds rows as needed. One per drive after the first (usually c:)

    By default, each new row gets a copy of all the data in the non Volume columns. This allows
    for filtering with Where-Object.

    Use the -ReportMode switch to provision a new empty row for all added rows. This is more readable.
.EXAMPLE 
    Get-MyServerCollection | Test-ServerConnectionOnPipeline | Get-ProcInfoOnPipeline | Get-VolumeInfoOnPipeline -ReportMOde | Select ComputerName,TotalProcs,ProcName,Cores,Volumes,DriveType,Capacity,PctFree | ft -autosize
    A Sampling of Functions and Columns in ReportMode 
.EXAMPLE
    Get-MyServerCollection | Test-ServerConnectionOnPipeline | Get-ProcInfoOnPipeline | Get-VolumeInfoOnPipeline | Select ComputerName,TotalProcs,Cores,Volumes,DriveType,Capacity,PctFree | Where PctFree -gt 95 | ft -autosize
    Gets all of the drives either over or under a threshold. IN this case % free greater than 95%
.EXAMPLE
    Get-MyServerCollection | Test-ServerConnectionOnPipeline | Get-OSCaptionOnPipeline | Get-TimeZoneOnPipeline | Get-TotalMemoryOnPipeline | Get-MachineModelOnPipeline | Get-ProcInfoOnPipeline | Get-VolumeInfoOnPipeline -ReportMOde | Select ComputerName,OSVersion,TotalMemory,MachineModel,TotalProcs,ProcName,Cores,Volumes,DriveType,Capacity,PctFree | Where DriveType -eq 3 | Export-Csv -path .\ServerSpecs.csv -NoTypeInformation
    Gets a nice specs report and outputs it to a CSV file in the working directory
#>
Function Get-VolumeInfoOnPipeline
    {
    [CmdletBinding()]

    Param
        (
        [parameter(
        Mandatory=$true,
        ValueFromPipeline= $true)]
        $ComputerProperties,

        [switch]
        $ReportMode
        )
    
    Begin
        {}
    Process
        {
        $ComputerProperties | Select *, Volumes, DriveType, Capacity, PctFree | %{
            If ((($_.Ping) -and ($_.WMI)) -or ($NoErrorCheck)) 
                {
                # When there is only one drive $Volumes.Count will be null
                # Otherwise we will use it to determine how many rows to add
                $Volumes = Get-WmiObject -computername $_.ComputerName -class Win32_Volume
                If ($Volumes.Count -eq $null) 
                    {
                    $_.Volumes = $Volumes.DriveLetter
                    $_.DriveType = $Volumes.DriveType
                    $_.Capacity = [Math]::Round(($Volumes.Capacity / 1GB), 0)
                    $_.PctFree = [Math]::Round($Volumes.FreeSpace/$Volumes.Capacity*100,1)
                    $_
                    If ($ReportMode){""}
                    New-Object PSObject -Property @{}
                    }
                # There is more than one drive
                Else 
                    {
                    $Count = $Volumes.Count - 1
                    For ($i=0; $i -le $Count; $i++) 
                        {
                        # For the first drive just fill in the normal row with the [0] (first) Value
                        If ($i -eq 0) 
                            {
                            $_.Volumes = $Volumes[$i].DriveLetter 
                            $_.DriveType = $Volumes[$i].DriveType
                            $_.Capacity = [Math]::Round(($Volumes[$i].Capacity / 1GB), 0)
                            $_.PctFree = [Math]::Round($Volumes[$i].FreeSpace/$Volumes[$i].Capacity*100,1)
                            $_
                            }
                        # For all subsequent rows a new (blank or copied) row must be created
                        Else 
                            {
                            # Calculate PctFree
                            $PctFree = ""
                            If ($Volumes[$i].DriveType -eq 3) {$PctFree = [Math]::Round($Volumes[$i].FreeSpace/$Volumes[$i].Capacity*100,1)}

                            If ($ReportMode)
                                {
                                # Here a brand new row is built without all of the other deatils
                                New-Object PSObject -Property @{
                                    # ComputerName = $_.ComputerName
                                    Volumes = $Volumes[$i].DriveLetter
                                    Capacity = [Math]::Round(($Volumes[$i].Capacity / 1GB), 0)
                                    DriveType = $Volumes[$i].DriveType
                                    PctFree = $PctFree
                                    }
                                }
                            Else
                                {
                                # In this case a new row is built from a copy of the current row
                                # This preserves the full details
                                $_.Volumes = $Volumes[$i].DriveLetter 
                                $_.Capacity = [Math]::Round(($Volumes[$i].Capacity / 1GB), 0)
                                $_.DriveType = $Volumes[$i].DriveType
                                $_.PctFree = $PctFree
                                New-Object PsObject $_                                
                                }
                            }
                        } # End Loop
                        If ($ReportMode){""}
                        New-Object PSObject -Property @{}
                    }
                }
            Else
                {
                $_.Volumes = 'No Try'
                $_.Capacity = 'No Try'
                $_.DriveType = 'No Try'
                $_.PctFree = 'No Try'
                $_
                }
            }
        }
    }

<#
All Columns So Far
    Get-MyServerCollection
        ComputerName
        Role
        Location
    Test-ServerConnectionOnPipeline
        Ping
        WMI
        PSRemote
        BootTime
    Get-OSCaptionOnPipe
        OSVersion
    Get-TimeZoneOnPipeline
        TimeZone
    Get-TotalMemoryOnPipe
        TotalMemory
    Get-MachineModelOnPipeline
        MachineModel
    Get-ProcInfoOnPipe
        TotalProcs
        ProcName
        Cores
        DataWidth
#>


