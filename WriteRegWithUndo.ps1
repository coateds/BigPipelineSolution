Function Get-RegItemDetails
    {
    Param($Path)

    $regKey=Get-Item $Path
    $regkey.GetvalueNames() | % {New-Object PSObject -Property @{'Name' = $_; 'Type' = $regkey.getvaluekind($_); 'Value' = $regkey.getvalue($_)}  }
    }

Function Get-RegItemType
    {
    Param($Path,$Name)

    $regKey=Get-Item $Path
    ($regkey.GetvalueNames() | % {New-Object PSObject -Property @{'Name' = $_; 'Type' = $regkey.getvaluekind($_)}  } | Where Name -eq $Name).Type
    }

<#
$RegValue = Get-RegItemDetails -Path HKLM:\software\mysoftware | Where Name -eq MyNewItem5
If ($RegValue -eq $null) {"Not Found"}


$Type = Get-RegItemType -Path HKLM:\software\mysoftware -Name MyNewItem5
If ($Type -eq $null) {$Type = "Not Found"}
$Type
#>

<#
.EXAMPLE
    ('Server1','Server2') | Get-ServerObjectCollection | Test-ServerConnectionOnPipeline | 
    Set-RegItemWithUndo -Path HKLM:\software\mysoftware -Name MyNewItem -Type String -Value Fubart | 
    Select ComputerName,Path,Name,OriginalType,OriginalValue,Result  | ft

    Use: Invoke-Command -ScriptBlock {New-Item -Path HKLM:\software\MySoftware} -ComputerName (Get-MyServerCollection).ComputerName
    From the console to create the Reg Keys on all computers
.EXAMPLE
    Get-MyServerCollection | Test-ServerConnectionOnPipeline | 
    Set-RegItemWithUndo -Path HKLM:\software\mysoftware -Name MyNewItem -Type String -Value Mickey | 
    Select ComputerName,Path,Name,OriginalType,OriginalValue,Result  | Export-Csv .\RegMyNewItem.csv
#>
Function Set-RegItemWithUndo
    {
    [CmdletBinding()]

    Param
        (
        [parameter(
        Mandatory=$true, 
        ValueFromPipeline= $true)]
        $ComputerProperties,

        [switch]
        $NoErrorCheck,

        $Path,
        
        $Name,
        
        [ValidateSet('String', 'ExpandString', 'Binary', 'DWord', 'MultiString', 'QWord')] 
        $Type,
        
        $Value
        )
    
    Begin
        {}
    Process
        {
        $ComputerProperties | Select *, Path, Name, OriginalType, OriginalValue, Result | %{
            If ((($_.Ping) -and ($_.PSRemote)) -or ($NoErrorCheck))
                {
                # Write the Path and Name for the record
                $_.Path = $Path
                $_.Name = $Name

                $sb = 
                    {
                    $Path = $Args[0]
                    $Name = $Args[1]
                    
                    $regkey1 = (Get-Item $Path)
                    $regkey1.GetvalueNames() | 
                        % {New-Object PSObject -Property @{'Name' = $_; 'Type' = $regkey1.getvaluekind($_); 'Value' = $regkey1.getvalue($_)}} |
                        Where Name -eq $Name

                    # At this point, I need to add a success/fail column to the output object
                    }
                
                # Run the Script Block
                $RegValue = Invoke-Command -ComputerName $_.ComputerName -ScriptBlock $sb -ArgumentList $Path, $Name -ErrorAction SilentlyContinue
                
                $_.OriginalValue = $RegValue.Value
                $_.OriginalType = $RegValue.Type
                If ($_.OriginalValue -eq $null) {$_.OriginalValue = 'Not Found'}

                # Make another Invoke-Command call here with return of Success or Failure
                $sb1 = 
                    {
                    $Path = $Args[0]
                    $Name = $Args[1]
                    $Type = $Args[2]
                    $Value = $Args[3]

                    Try
                        {Set-ItemProperty -Path $Path -Name $Name -Type $Type -Value $Value -ErrorAction Stop}
                    Catch
                        {'Error'}
                    }

                $_.Result = Invoke-Command -ComputerName $_.ComputerName -ScriptBlock $sb1 -ArgumentList $Path, $Name, $Type, $Value
                If ($_.Result -eq $null) {$_.Result = 'Success'}

                # Now write the whole mess out to the output
                $_
                }
            }
        }
    }

Function Set-UndoRegSettingFromCsv
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
        $ComputerProperties | Select *, UndoResult | %{
            If ((($_.Ping) -and ($_.PSRemote)) -or ($NoErrorCheck))
                {
                'For ' + $_.ComputerName + ' -'
                '    Depending on Result: ' + $_.Result
                '    Reset ' + $_.Path + ' - ' + $_.Name
                '    To ' + $_.OriginalType + ': ' + $_.OriginalValue
                }
            # $_
            }
        }
    }

    # $Path = 'HKLM:\software\mysoftware'
    # $Name =  'MyNewItem'
    # $regkey1 = (Get-Item $Path)
    # $regkey1.GetvalueNames() | % {New-Object PSObject -Property @{'Name' = $_; 'Type' = $regkey1.getvaluekind($_); 'Value' = $regkey1.getvalue($_)}  }