#requires -version 2.0
# handy global variables

$GLOBAL:SMADLL   = ([appdomain]::CurrentDomain.getassemblies()|?{$_.location -match "System.Management.Automation.dll"}).location
$GLOBAL:DATAGENDIR = "$psScriptRoot\DataGen"

# now add the PSScriptRoot to the path, this will ensure that any scripts
# in the module directory are accessible
$env:path += ";$psscriptroot;$psscriptroot\Scripts"
[System.AppDomain]::CurrentDomain.SetData("APP_CONFIG_FILE", (join-path $psScriptRoot cmdblets.dll.config))

function Get-CMDBCommand
{
    [CmdletBinding()]
    param ( )
    end
    {
        get-command -module CMDBLets| sort-object CommandType,Noun
    }
}

$CMDBLetsTypesFile = join-path $PSScriptRoot CMDBLets.Types.ps1xml
#update-typedata $CMDBLetsTypesFile -ErrorAction SilentlyContinue
$CMDBLetsFormatFile = join-path $PSScriptRoot CMDBLets.Format.ps1xml
update-formatdata $CMDBLetsFormatFile -ErrorAction SilentlyContinue