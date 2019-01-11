<#
.SYNOPSIS
    PowerShell script that adds Ola Hallengren's SQL maintenance scripts to a MS SQL Server hosting SCCM databases.
    Script also configures the SQL Agent to run the maintenance jobs on a schedule.
.DESCRIPTION
    Designed to run as an SCCM Configuration Item deployed to all SCCM SUP and SQL/Database servers.
    - NOTE: Create dynamic collections in SCCM for 'All SCCM SUP servers' and 'All SCCM SQL Servers' and deploy the baseline to a new collection with the aforementioned collections as 'includes'.
    The script will run as the SYSTEM account, and therefore should have access to the SQL server in order to create the Stored Procedures and SQL Agent Jobs+Schedules.
    The MaintenanceSolution.sql file from Ola Hallengren's site should be placed on the network so that the SCCM Primary Site servers can access it.
    Replacing the $OlaHallengrenScriptLocation location with the latest MaintenanceSolution.sql will automatically upgrade the scripts on the SQL server.
.NOTES
    Return Codes:
    - All output including errors are logged to $LogFile

    title     : SCCM SQL Ola Hallengren Maintenance Scripts.ps1
    Author    : Jack Rudlin
    Change History:

     Date           Author           Version   Comments
     18 Oct 2018    Jack Rudlin      1.0       Created Script
     24 Oct 2018    Jack Rudlin      1.1       Updated following script review by CR
     03 Nov 2018    Jack Rudlin      1.2       Changed User Databases schedule to daily instead of weekly. Added custom step mods property.
#>


#region #############################################################################################
Try {

    # Variables
    $ErrorActionPreference = "Stop"

    $OlaHallengrenScriptLocation = "\\domain.com\Admin\SQL\Ola Hallengren\MaintenanceSolution.sql" # NOTE: The SYSTEM/Computer account of the site database server will need access to this location
    If(-not(Test-Path -Path $OlaHallengrenScriptLocation)){write-output -InputObject "Could not access or find the sql script @ $OlaHallengrenScriptLocation. Please check this location exists";return}

    $OlaScriptVersionDetection = "Version: "

    $SQLServerPoShModule = "\\domain.com\Admin\SQL\public\PSModules\SqlServer\21.0.17279\SqlServer.psm1" # SQLServer PoSh module

    $SQLInstance = $env:COMPUTERNAME

    $LogFile = "c:\admin\Log\SCCM\SQL\sccm_sql_olahallengren.log"
    $component = "SCCM SQL Ola Hallengren Maintenance Scripts"

    # Ola Hallengren SQL Agent Jobs # NOTE: don't deviate from the format of days/times required by Microsoft.SqlServer.Management.Smo.Agent.JobSchedule
    $SQLJobsConfig = @(
        [pscustomobject]@{
            Name="CommandLog Cleanup";
            ScheduleType="Weekly";
            ScheduleDay=64;
            ScheduleTime="00:00:00";
        }
        [pscustomobject]@{
            Name="DatabaseBackup - SYSTEM_DATABASES - FULL";
            Enabled=$false;
        }
        [pscustomobject]@{
            Name="DatabaseBackup - USER_DATABASES - DIFF";
            Enabled=$false;
        }
        [pscustomobject]@{
            Name="DatabaseBackup - USER_DATABASES - FULL";
            Enabled=$false;
        }
        [pscustomobject]@{
            Name="DatabaseBackup - USER_DATABASES - LOG";
            Enabled=$false;
        }
        [pscustomobject]@{
            Name="DatabaseIntegrityCheck - SYSTEM_DATABASES";
            ScheduleType="Weekly";
            ScheduleDay=64;
            ScheduleTime="00:15:00";
        }
        [pscustomobject]@{
            Name="DatabaseIntegrityCheck - USER_DATABASES";
            ScheduleType="Weekly";
            ScheduleDay=64;
            ScheduleTime="00:30:00";
        }
        [pscustomobject]@{
            Name="IndexOptimize - USER_DATABASES";
            CustomStepMod="@UpdateStatistics = 'ALL'"
            ScheduleType="Weekly";
            ScheduleDay=1;
            ScheduleTime="00:30:00";
        }
        [pscustomobject]@{
            Name="Output File Cleanup";
            ScheduleType="Weekly";
            ScheduleDay=64;
            ScheduleTime="00:05:00";
        }
        [pscustomobject]@{
            Name="sp_delete_backuphistory";
            Enabled=$false;
        }
        [pscustomobject]@{
            Name="sp_purge_jobhistory";
            ScheduleType="Weekly";
            ScheduleDay=64;
            ScheduleTime="00:10:00";
        }
    )

#endregion #############################################################################################

#region #############################################################################################

    # Test to see if the SqlServer module is loaded, and if not, load it
    if (-not(Get-Module -name 'SqlServer')) {
       Import-Module -Name $SQLServerPoShModule
    }

    if (-not(Get-Module -name 'SqlServer')) {

        Write-Output -InputObject "Could not load SqlServer module, check the SqlServer is available @ $SQLServerPoShModule"
        return

    }

#endregion #############################################################################################

Function Add-TextToCMLog {
    <#
    .SYNOPSIS
    Log to a file in a format that can be read by Trace32.exe / CMTrace.exe

    .DESCRIPTION
    Write a line of data to a script log file in a format that can be parsed by Trace32.exe / CMTrace.exe

    The severity of the logged line can be set as:

            1 - Information
            2 - Warning
            3 - Error

    Warnings will be highlighted in yellow. Errors are highlighted in red.

    The tools to view the log:

    SMS Trace - http://www.microsoft.com/en-us/download/details.aspx?id=18153
    CM Trace - Installation directory on Configuration Manager 2012 Site Server - <Install Directory>\tools\

    .EXAMPLE
    Add-TextToCMLog c:\output\update.log "Application of MS15-031 failed" Apply_Patch 3

    This will write a line to the update.log file in c:\output stating that "Application of MS15-031 failed".
    The source component will be Apply_Patch and the line will be highlighted in red as it is an error
    (severity - 3).

    #>

        #Define and validate parameters
        [CmdletBinding()]
        Param(
            #Path to the log file
            [parameter(Mandatory=$True)]
            [String]$LogFile,

            #The information to log
            [parameter(Mandatory=$True)]
            [String]$Value,

            #The source of the error
            [parameter(Mandatory=$True)]
            [String]$Component,

            #The severity (1 - Information, 2- Warning, 3 - Error)
            [parameter(Mandatory=$True)]
            [ValidateRange(1,3)]
            [Single]$Severity
            )


        #Obtain UTC offset
        $DateTime = New-Object -ComObject WbemScripting.SWbemDateTime
        $DateTime.SetVarDate($(Get-Date))
        $UtcValue = $DateTime.Value
        $UtcOffset = $UtcValue.Substring(21, $UtcValue.Length - 21)

        # Delete large log file
        If(test-path -Path $LogFile -ErrorAction SilentlyContinue)
        {
            $LogFileDetails = Get-ChildItem -Path $LogFile
            If ( $LogFileDetails.Length -gt 5mb )
            {
                Remove-item -Path $LogFile -Force -Confirm:$false
            }
        }
        else
        {
            new-item -Path ($LogFile | Split-Path -Parent) -Force -Confirm:$false -ItemType Directory | Out-Null
        }

        #Create the line to be logged
        $LogLine =  "<![LOG[$Value]LOG]!>" +`
                    "<time=`"$(Get-Date -Format HH:mm:ss.fff)$($UtcOffset)`" " +`
                    "date=`"$(Get-Date -Format M-d-yyyy)`" " +`
                    "component=`"$Component`" " +`
                    "context=`"$([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " +`
                    "type=`"$Severity`" " +`
                    "thread=`"$($pid)`" " +`
                    "file=`"`">"

        #Write the line to the passed log file
        Out-File -InputObject $LogLine -Append -NoClobber -Encoding Default -FilePath $LogFile -WhatIf:$False

        Switch ($component) {

            1 { Write-Information -MessageData $Value }
            2 { Write-Warning -Message $Value }
            3 { Write-Error -Message $Value }

        }

        write-output -InputObject $Value

    }

Function Install-OlaSQLMaintenance
{
<#
.DESCRIPTION
   Invoke the SQL CMD that installs the Ola Hallengren SQL maintenance stored procedures into the Master database
#>
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$True,Position=1)]
    $ScriptContent,
    [Parameter(Mandatory=$True)]
    [string]$SQLInstance,
    [Parameter(Mandatory=$false)]
    $UpgradeExistingScript = $false
    )


    If($UpgradeExistingScript){

        Add-TextToCMLog -LogFile $LogFile -Value "`nUpgradeExistingScript set to True. Setting 'SET @CreateJobs = 'N'' " -Component $component -Severity 2
        $ScriptContent = $ScriptContent | ForEach-Object {If($_ -like 'SET @CreateJobs*'){$_.Replace('Y','N')} else {$_}} -ErrorAction Stop
        Add-TextToCMLog -LogFile $LogFile -Value "Script succesfully adjusted" -Component $component -Severity 1
    }

    Add-TextToCMLog -LogFile $LogFile -Value "`nRunning Ola Hallengren maintenance script installation...." -Component $component -Severity 1
    $ScriptContent = $ScriptContent | Out-String
    Invoke-Sqlcmd -ServerInstance $SQLInstance -Query $ScriptContent -AbortOnError

}



$OlaHallengrenScriptcontent = Get-Content -Path $OlaHallengrenScriptLocation -ErrorAction Stop
Add-TextToCMLog -LogFile $LogFile -Value "Ola Hallengren script contents imported into variable from location: $OlaHallengrenScriptLocation" -Component $component -Severity 1
$AgentService = Get-Service -Name SQLSERVERAGENT -ErrorAction SilentlyContinue

If(-not($AgentService)){

    Add-TextToCMLog -LogFile $LogFile -Value "Could not find SQL Agent service on this server" -Component $component -Severity 3
    return

}

else

{

    Add-TextToCMLog -LogFile $LogFile -Value "SQL Agent service found" -Component $component -Severity 1

    If($AgentService.StartType -ne "automatic"){

        $AgentService | Set-Service -StartupType Automatic
        Add-TextToCMLog -LogFile $LogFile -Value "Setting SQL Agent service has now been set startup to automatic" -Component $component -Severity 2

    } else {
        Add-TextToCMLog -LogFile $LogFile -Value "SQL Agent service already set to startup automatically" -Component $component -Severity 1
    }

    Add-TextToCMLog -LogFile $LogFile -Value "SQL Agent service status is: $($AgentService.Status)" -Component $component -Severity 1

    If($AgentService.Status -ne "Running"){
        Add-TextToCMLog -LogFile $LogFile -Value "Starting SQL Agent service" -Component $component -Severity 1
        $AgentService | Start-Service -Confirm:$false
        Start-Sleep 5
        If((Get-Service -Name SQLSERVERAGENT).Status -ne "Running"){
            Add-TextToCMLog -LogFile $LogFile -Value "Could not successfully start SQL Agent service" -Component $component -Severity 3
            return
        }
    }
    

}

# Version check any existing Ola maintenance script
[string]$TSQL_CheckOlaVersion = "
DECLARE @VersionKeyword nvarchar(max)

SET @VersionKeyword = '--// Version: '

SELECT sys.schemas.[name] AS SchemaName,
       sys.objects.[name] AS ObjectName,
       CASE WHEN CHARINDEX(@VersionKeyword,OBJECT_DEFINITION(sys.objects.[object_id])) > 0 THEN SUBSTRING(OBJECT_DEFINITION(sys.objects.[object_id]),CHARINDEX(@VersionKeyword,OBJECT_DEFINITION(sys.objects.[object_id])) + LEN(@VersionKeyword) + 1, 19) END AS [Version],
       CAST(CHECKSUM(CAST(OBJECT_DEFINITION(sys.objects.[object_id]) AS nvarchar(max)) COLLATE SQL_Latin1_General_CP1_CI_AS) AS bigint) AS [Checksum]
FROM sys.objects
INNER JOIN sys.schemas ON sys.objects.[schema_id] = sys.schemas.[schema_id]
WHERE sys.schemas.[name] = 'dbo'
AND sys.objects.[name] IN('CommandExecute','DatabaseBackup','DatabaseIntegrityCheck','IndexOptimize')
ORDER BY sys.schemas.[name] ASC, sys.objects.[name] ASC
"

$InstalledVersion = Invoke-Sqlcmd -ServerInstance $SQLInstance -Query $TSQL_CheckOlaVersion

# Version check the maintenance script file
$OlaScriptVersion = ($OlaHallengrenScriptcontent | Where-Object -FilterScript {$_ -like "$OlaScriptVersionDetection*"}).TrimStart($OlaScriptVersionDetection)

# Check if the maintenance scripts are already installed and match installed version if so
If(-not($InstalledVersion))
{

    Add-TextToCMLog -LogFile $LogFile -Value "`nMaintenance scripts not currently installed on this SQL server. They will be installed." -Component $component -Severity 2
    Install-OlaSQLMaintenance -ScriptContent $OlaHallengrenScriptcontent -SQLInstance $SQLInstance

}
ElseIf([datetime]$InstalledVersion.Version[0] -eq [datetime]$OlaScriptVersion)
{
    Add-TextToCMLog -LogFile $LogFile -Value "`nScripts installed and the version installed matches the version at: $OlaHallengrenScriptLocation" -Component $component -Severity 1
}
elseif([datetime]$InstalledVersion.Version[0] -lt [datetime]$OlaScriptVersion)
{
    Add-TextToCMLog -LogFile $LogFile -Value "Scripts will be upgraded with the new Ola maintenance scripts version: $OlaScriptVersion" -Component $component -Severity 2
    Install-OlaSQLMaintenance -ScriptContent $OlaHallengrenScriptcontent -SQLInstance $SQLInstance -UpgradeExistingScript $true
}
elseif([datetime]$InstalledVersion.Version[0] -gt [datetime]$OlaScriptVersion)
{
    Add-TextToCMLog -LogFile $LogFile -Value "`nThe SQL server $SQLInstance has been updated to a newer version than in the file: $OlaHallengrenScriptLocation" -Component $component -Severity 2
    Add-TextToCMLog -LogFile $LogFile -Value "Please download the latest script file from https://github.com/olahallengren/sql-server-maintenance-solution/blob/master/MaintenanceSolution.sql and overwrite the file at: $OlaHallengrenScriptLocation" -Component $component -Severity 2
}

Add-TextToCMLog -LogFile $LogFile -Value "`nNow checking SQL Agent Jobs configuratoin and schedule..." -Component $component -Severity 1
ForEach($JobConfig in $SQLJobsConfig){

    # Search for SQL Agent Job to amend the schedule for
    $CurrentJob = Get-SqlAgentJob -ServerInstance $SQLInstance -Name $JobConfig.Name

    # Quit if no Job found
    if ($CurrentJob -eq $Null)
    {
        Add-TextToCMLog -LogFile $LogFile -Value "No Job with that name: $($JobConfig.Name)" -Component $component -Severity 2
        return
    }

    # Set the job to enabled or disabled depending on the setting in the jobconfig hash table
    If($JobConfig.Enabled -eq $false)
    {
        $CurrentJob.IsEnabled = $false
    }
    else
    {
        $CurrentJob.IsEnabled = $true
    }

    # Alter the job with the IsEnable property set above
    $CurrentJob.Alter()

    # Only process jobs that need schedules
    If($JobConfig.ScheduleType){

        # Check if a schedule already exists
        $Schedule = $null

        # Create schedule name
        $JobScheduleName = "$($JobConfig.ScheduleType)-$($JobConfig.Name)"

        # Get current schedule if it exists
        $Schedule = $CurrentJob | Get-SqlAgentJobSchedule -Name $JobScheduleName -ErrorAction SilentlyContinue

        # Determine create new schedule or alter existing one
        If($Schedule)
        {
            Add-TextToCMLog -LogFile $LogFile -Value "`nSchedule '$JobScheduleName' found. Setting job schedule properties..." -Component $component -Severity 1
            $WriteJobScheduleType = "Alter"
        }
        else
        {
            Add-TextToCMLog -LogFile $LogFile -Value "`nCould not find an existing job schedule named '$JobScheduleName' for job: '$($JobConfig.Name)'" -Component $component -Severity 2
            Add-TextToCMLog -LogFile $LogFile -Value "Will now create new schedule..." -Component $component -Severity 1
            $Schedule = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Agent.JobSchedule -ArgumentList ($CurrentJob, $JobScheduleName)
            $WriteJobScheduleType = "Create"
        }

        # Build and implement schedule
        If($Schedule){
            $Schedule.ActiveEndDate = Get-Date -Month 12 -Day 31 -Year 9999
            $Schedule.ActiveEndTimeOfDay = '23:59:59'
            $Schedule.FrequencyTypes = $JobConfig.ScheduleType
            $Schedule.FrequencyRecurrenceFactor = 1
            $Schedule.FrequencySubDayTypes = "Once"
            $Schedule.FrequencyInterval = $JobConfig.ScheduleDay
            $Schedule.ActiveStartDate = Get-Date
            $Schedule.ActiveStartTimeOfDay = $JobConfig.ScheduleTime
            $Schedule.IsEnabled = $true
            $Schedule.$($WriteJobScheduleType)()
        }

        If(Get-SqlAgentJobSchedule -Name $JobScheduleName -InputObject $CurrentJob){
            Add-TextToCMLog -LogFile $LogFile -Value "Schedule '$JobScheduleName' for job: '$($JobConfig.Name)' created" -Component $component -Severity 1
        } else {
            Add-TextToCMLog -LogFile $LogFile -Value "Could not create SQL Agent job schedule '$JobScheduleName' for job: '$($JobConfig.Name)'" -Component $component -Severity 2
        }

    } else {

        Add-TextToCMLog -LogFile $LogFile -Value "`nNo schedule specified for this Job: '$($JobConfig.Name)' so will not create any schedule" -Component $component -Severity 1

    }

    # Custom step modification
    If($JobConfig.CustomStepMod)
    {
        Add-TextToCMLog -LogFile $LogFile -Value "Custom step modification specified for this Job: '$($JobConfig.Name)' so will add $($JobConfig.CustomStepMod)" -Component $component -Severity 1
        $Step = $CurrentJob.JobSteps | Where-Object -Property Name -EQ $CurrentJob.Name
        If($Step.count -eq 1)
        {
            Add-TextToCMLog -LogFile $LogFile -Value "Found step: '$($Step.Name)' so will add: '$($JobConfig.CustomStepMod)'" -Component $component -Severity 1
            If($CurrentJob.JobSteps.Command -match $JobConfig.CustomStepMod)
            {
                Add-TextToCMLog -LogFile $LogFile -Value "Custom step config: '$($JobConfig.CustomStepMod)' already added." -Component $component -Severity 1
            }
            else
            {
                Add-TextToCMLog -LogFile $LogFile -Value "Adding step config: '$($JobConfig.CustomStepMod)' to step now...." -Component $component -Severity 1
                $Step.Command = $Step.Command + ','
                $Step.Command = $Step.Command + "`n"
                $Step.Command = $Step.Command + $JobConfig.CustomStepMod
                $Step.Alter()
            }
        }
        else
        {
            Add-TextToCMLog -LogFile $LogFile -Value "Couldn't find step: '$($CurrentJob.Name)' to modify, so will NOT add: '$($JobConfig.CustomStepMod)'" -Component $component -Severity 2
        }
        


    }

}

Add-TextToCMLog -LogFile $LogFile -Value "Script finished" -Component $component -Severity 1

} Catch {
    write-output -InputObject "$component encountered an error. Please check the log file $logfile"
    Add-TextToCMLog -LogFile $LogFile -Value "$component encountered an error. Please check the log file $logfile" -Component $component -Severity 3
}
