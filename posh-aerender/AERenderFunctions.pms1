
function Get-AEInstall {
    [CmdletBinding()]
    param()
    begin {
        $RegPaths = Join-Path -Path @(  'HKLM:\SOFTWARE', 'HKLM:\SOFTWARE\WOW6432Node' ) -ChildPath 'Microsoft\Windows\CurrentVersion\Uninstall'
        $Properties = @(
            'InstallLocation', 'InstallDate', 'DisplayName',
            @{ Name = 'DisplayVersion'   ; Expression = { [version]::new( $RegKey.GetValue('DisplayVersion') ) ?? $null } },
            @{ Name = 'RegistryReference'; Expression = { $RegValues.PSPATH ?? $null } }
        )
    }
    process {
        $RegPaths | Get-ChildItem -PipelineVariable RegKey | Where-Object { $RegKey.GetValue('DisplayName') -ILike '*After*Effects*' } | Get-ItemProperty -PipelineVariable RegValues |
            Select-Object $Properties | Sort-Object DisplayVersion -Descending
    }
}

function Get-AERenderPath {
    [CmdletBinding()]
    param()
    process {
        Get-AEInstall | Write-Output -PipelineVariable AEInstall | ForEach-Object {
            $AEInstall | Add-Member -NotePropertyName 'AERenderPath' -NotePropertyValue (
                (Get-ChildItem -Path $_.InstallLocation -Filter '*After*Effects*' | Get-ChildItem -Filter 'Support Files' | Get-ChildItem -Filter 'aerender.*').FullName ?? $null
            )  -PassThru | Where-Object { $null -ne $_.AERenderPath } | Select-Object -First 1 -ExpandProperty AERenderPath
        }
    }
}

function Invoke-AERender {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)] [string] [Alias('Project','PSPath','LP')] $ProjectPath,
        [Parameter()] [int] $StartFrame,
        [Parameter()] [int] $EndFrame,
        [Parameter()] [int] $Incrememnt,
        [Parameter()] [string] $AERenderPath = (Get-AERenderPath),
        [Parameter(ValueFromRemainingArguments)] [string[]] $AdditionalArguments
    )
    begin {
        $ArgumentNames = @{
            StartFrame = '-start_frame'
            EndFrame   = '-end_frame'
            Increment  = '-increment'
        }
    }
    process {
        # Create list to collect arguments to pass to aerender.exe and add the required -project argument
        $AERenderArgs = [System.Collections.Generic.List[string]]::new( [string[]] (  '-project', "$ProjectPath" ) )
        # Add required -project argument
        # $AERenderParams.AddRange( [string[]] (  '-project', "$ProjectPath" ) )
        # Add -start_frame, -end_frame, and -increment arguments if they have been provided
        $ArgumentNames.GetEnumerator() | Where-Object { $PSBoundParameters.ContainsKey( $_.Key ) } -PipelineVariable Argument | ForEach-Object {
            $AERenderArgs.AddRange( [string[]] ( $Argument.Value, $PSBoundParameters[ $Argument.Key ] ) )
        }
        $ArgumentNames.GetEnumerator().Where{ $PSBoundParameters.ContainsKey( $_.Key ) }.ForEach{ $AERenderArgs.AddRange( [string[]] ( $_.Value, $PSBoundParameters[ $_.Key ] ) ) }
        $AERenderArgs.AddRange( $AdditionalArguments )
        & $AERenderPath $AERenderArgs
    }
}
