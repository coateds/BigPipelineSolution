<#
BigPipeLineSolution PSRemote Session Branch

This variant: 
1. Opens one PSSession connection to a remote computer
2. Stores the session object as a property in the $ComputerProperties object collection
3. Uses this object as needed in invoke-command
4. closes the session in one last function on the pipe:

Example: Get-MyServerCollection | Test-ServerConnectionOnPipeline | Get-TimeZoneOnPipeline | Get-CultureOnPipeline | Cleanup-PSSession | ft


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
.Synopsis 
    Runs availability checks on servers
.DESCRIPTION 
    This typically takes an imported csv file with a ComputerName Column as an imput object
    but just about any collection of objects that exposes .ComputerName should work
    The output is the same type of object as the input (hopefully) so that it can be piped 
    to the next function to add another column.

    I am changing this function to attempt to open a new-pssession
    If successful, store the Session object in the computerProperties object being passed on the pipeline
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
        $ComputerProperties | Select *, Ping, WMI, PSRemote, PSSession, BootTime | %{
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

                $Session = New-PSSession -ComputerName $ComputerProperties.ComputerName -ErrorAction SilentlyContinue

                If ($Session -ne $null) 
                    {$_.PSRemote = $true; $_.PSSession = $Session}
                Else
                    {$_.PSRemote = $false}
                }
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

    Modifying to make Invoke-Command use existing $PSSession object to prevent from having to open a new connection
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
                $_.TimeZone = Invoke-Command -Session $_.PSSession -ScriptBlock $sb
                }
            Else {$_.TimeZone = 'No Try'}
            $_
            }
        }
    }

Function Get-CultureOnPipeline
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
        $ComputerProperties | Select *, Culture | %{
            If ((($_.Ping) -and ($_.PSRemote)) -or ($NoErrorCheck))
                {
                $_.Culture = $null
                
                $sb = {(Get-Culture).Name}
                $_.Culture = Invoke-Command -Session $_.PSSession -ScriptBlock $sb
                }
            Else {$_.Culture = 'No Try'}
            $_
            }
        }
    }

Function Cleanup-PSSession
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
        {Remove-PSSession -Session $_.PSSession; $_}
    }