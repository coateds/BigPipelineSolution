# BigPipelineSolution

Despite knowing that the pipeline is supposed to be the power behind PowerShell, I have struggled with it. Like everyone else, I have seen the silly pipeline tricks that whack rogue instances of Notepad as if that basic editor were the great scripting enemy. Recently, I started a contract where I was 

handed a listing of servers that were mine to administer and look after. There were categories within the list and I realized that by converting it to a 

CSV file, I could use Import-Csv to get an instant list of servers or even a subset of servers. It occurred to me that this would be very useful as the 

first item on the left of a pipeline.

For compatibility with many existing cmdlets, the CSV file should store the server names in a column with the heading 'ComputerName'. That makes the 

following command possible to determine where a particular service is running. With this you can determine where a service is running in your list of 

servers and its status. It is unfortunate, but instructive, to note that it is the ComputerName property passed from the Import-Csv to Get-Service, yet 

MachineName is the property selected from Get-Service. Use Get-Help Get-Service -Full and Get-Service | Get-Member to see these values.

Import-Csv .\Servers.csv | Get-Service winrm | Select MachineName,status,name

BTW, if you are really determined to play network-wide Whac-A-Notepad, "Import-Csv .\Servers.csv | Get-Process Notepad..." will get you started down that 

path.

In my included example, BigPipelineSolution.ps1, I put a wrapper around Import-Csv .\Servers.csv. The Get-MyServerCollection function takes advantage of 

Intellisense to fill in possible filter values provided by the CSV file. In this case Role and Location. This would allow me to quickly limit the scope 

of the servers queried to 'SQL servers in Arizona' for instance. I started thinking about other columns that could be included in the CSV file when I 

realized I could add columns to the object being passed along the pipeline and suddenly a whole array of possibilities came to mind. 

To begin with, a series of information gathering functions could add column after column. Using WMI, for instance, each function could add columns for 

process, memory, operating system etc. However, it would be ineffiecient to keep querying a server that could not be contacted. So the second function, 

number 2 on the left of the pipeline, should verify server health. In my example, this function is Test-ServerConnectionOnPipeline.

The function adds three columns to the objects on the pipeline: Ping, WMI and PSRemote. Each is populated True or False depending on whether the server 

passes that particular test. The following command will run the three tests on all of the servers in the CSV file. As an added bonus, you get the last 

BootTime for each server as it is returned by the WMI query used to test WMIs availability.

Get-MyServerCollection | Test-ServerConnectionOnPipeline | ft -autosize

ComputerName Role Location Ping  WMI PSRemote BootTime             
------------ ---- -------- ----  --- -------- --------             
Server1      DC   WA       True True     True 12/28/2016 2:40:21 PM
Server2      Web  WA       True True     True 12/28/2016 2:40:22 PM
Server3      Web  AZ       True True     True 12/28/2016 2:40:23 PM
Server4      SQL  WA       True True     True 12/28/2016 2:40:23 PM
Server5      SQL  AZ       True True     True 12/28/2016 2:40:24 PM

Note that the output is just like any other collection of objects in PowerShell. The output can be piped to Select-Object, Where-Object and Export-Csv as 

well as all of the other standard cmdlets PS employs.

The example provided here includes the following functions and added columns:

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
    Get-VolumeInfoOnPipeline
        Volumes
        Capacity
        DriveType
        PctFree

So if you want a report on the OS version for all of your servers in Washington, use this:

Get-MyServerCollection -Location WA | Test-ServerConnectionOnPipeline | Get-OSCaptionOnPipeline | Select ComputerName, OSVersion | ft

Note that the last function, Get-VolumeInfoOnPipeline adds rows to the output in addition to columns. That is, there is a line for C:, D:, E: as 

required. This creates some interesting challenges and I leave it to you to decide if you like my solution or not.

Finally, there is a function, Get-ServerObjectCollection, that allows the conversion of a simple list of server names into a collection of objects as if 

from a CSV file with just one column heading, ComputerName. Use the following command to start the pipeline for an arbitrary list of servers. Examples 

are provided for the function to include ways to read a txt file and even use Active Directory.

('Server2','Server4') | Get-ServerObjectCollection | Test-ServerConnectionOnPipeline | ft

Take a very close look at this last function for some important clues to using the pipeline. First, in the Param block, the function accepts the whole 

object from the pipeline, but in this case it has specifically cast the object as a string. None of the other functions in this example explicitly cast 

the input object. Second, the Process block, which runs once for every object coming into the pipeline creates a generic object out of a hash table with 

just one key pair. This object is then passed from function to function along the pipeline. Hence the Param block for all of the other functions do not 

cast the incoming object. It is important to realize that the output of the function is a single object. It is only at the end that PowerShell outputs 

all of the objects as a collection of objects that can be assigned to a variable, output to the screen or exported as a CSV file.


    Param(
        [parameter(
        Mandatory=$true,
        ValueFromPipeline= $true)]
        [string]
        $ComputerName
    )


    Process
        {
        New-Object PSObject -Property @{'ComputerName' = $_}
        }


I think it is useful to think of the pipeline as a foreach loop with extra features. The following command illustrates this point. A Foreach cmdlet has a 

Begin, Process and End script block just like an advanced function.

(1..10) | % -Begin {'Begin'} -Process {$_} -End {'End'}

To add a column to each object the function has this basic process. 

    Process {$<ObjectFromPipeline> | Select *, NewColumnHeader | %{$_.NewColumnHeader='<ValueForTheCurrentRow>'; $_}}

    1. Select all of the properties of the current object plus one or more new ones
    2. Assign a value to each property of the current object
    3. Output the current object

The new properties appear as new columns when PowerShell assembles all of the objects into a collection at the end.

Conclusion and possible next steps. This started as an exercise in learning how to use the pipeline by scripting functions to be used on the pipeline. 

checkout this oneliner for a big capacity and configuration report on a list of servers. The biggest problem is adding new rows for servers that have 

more than one fixed drive. The extra rows are fine for just the one list of resources per server, but this will break down if you were to add another 

item for which there might be multiple rows. Getting IP address for servers that have more than one NIC for instance. Perhaps someone out there has an 

idea?

Get-MyServerCollection | Test-ServerConnectionOnPipeline | Get-OSCaptionOnPipeline | Get-TimeZoneOnPipeline | Get-TotalMemoryOnPipeline | Get-

MachineModelOnPipeline | Get-ProcInfoOnPipeline | Get-VolumeInfoOnPipeline -ReportMode | Select 

ComputerName,OSVersion,TotalMemory,MachineModel,TotalProcs,ProcName,Cores,Volumes,DriveType,Capacity,PctFree | Where DriveType -eq 3 | Export-Csv -path 

.\ServerSpecs.csv -NoTypeInformation

Another way to use this would be to log the results of taking some action on a list of servers that would be exported as a CSV. At a minimum this could 

be used to note the changes not made due to errors. It could also be used to create an 'undo' table. Consider the idea where a registry change is desired 

on a large number of servers. Start with a CSV file of the servers and add columns for existing value and type plus a success/fail column. The resulting 

collection of objects could be exported to a new CSV that could be used to remediate problems and even run an undo to set everything back to what it was 

originally.

