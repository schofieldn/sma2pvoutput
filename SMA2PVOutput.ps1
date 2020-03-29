# SMA2PVOutput v2.0
# Last updated 27/02/2015
# Written by Neil Schofield (neil.schofield@sky.com)
#
# This Powershell script provides a means of exporting live and historic PV data
# to PVOutput.org. Registering your PV system on PVoutput.org is a pre-requisite
# to using SMA2PVOutput - see README.TXT.
#
# The script has been tested with Powershell 4.0.
#
# Use Task Scheduler to run the script periodically during daylight hours, (say
# every 15 minutes at 1 minute past the quarter hour). An end-of-day run can be
# scheduled using -BatchMode to pick up any missing statuses.
#
# The script can use statuses obtained from an existing CSV export file, or it
# can invoke Sunny Explorer, if it is installed, to download a new one using the
# -ExportFromInverter switch.
#
# SYNTAX
#  .\SMA2PVOutput.ps1 [[-Date] yyyyMMdd] [-ExportFromInverter] [-BatchMode]
#
# For more information, see the accompanying README.TXT.
#
# *** This script is provided 'as is' without any warranty of any kind ***
#
# 
param (
[validatescript ({[datetime]::parseexact($_,"yyyyMMdd",$null) -le (Get-Date)})][string]$Date = (Get-Date -format yyyyMMdd),
[switch]$ExportFromInverter = $false,
[switch]$BatchMode = $false
)
# ==============================================================================
# The following values should be modified according to your system:

$SystemId = "nnnn"                                         # Specify your pvoutput.org System Id (in quotes).
$ApiKey = "ffffffffffffffffffffffffffffffffffffffff"       # Specify the API Key (in quotes) for your system on pvoutput.org.
$SEPath = "C:\Program Files (x86)\SMA\Sunny Explorer"      # Installation folder for Sunny Explorer, used to locate sunnyexplorer.exe.
$ExportPath = "C:\Users\Bob\Documents\SMA\Sunny Explorer"  # Location of daily files (CSV) exported from Sunny Explorer.
$PlantName = "Example Plant Name"                          # The name of your PV plant as it appears in Sunny Explorer.
$PlantExtension = "sx2"                                    # The extension of the file used by Sunny Explorer to store your plant information. Can be sx2 or sxp.
$UserPwd = "0000"                                          # The plant password for the "User" account. The default is 0000.

# The following values should generally not be changed:
$UploadLog = "$ExportPath\SMA2PVOutput.log"                # Used to provide diagnostic information and as a journal for previously uploaded statuses.
$ResponseLog = "$ExportPath\SMA2PVOutputResp.log"          # Output file containing responses from PVOutput web services.

# End of values
# ==============================================================================

$ErrorActionPreference = "Stop"

# What's the timestamp of the CSV before we start?
if (test-path "$ExportPath\$PlantName-$Date.csv")
{
    $OriginalCSVTime = [datetime](Get-ItemProperty -Path "$ExportPath\$PlantName-$Date.csv").lastwritetime
}
else
{
    $OriginalCSVTime = [datetime](Get-Date -hour 0 -minute 0 -second 0 -year 1 -month 1 -day 1)
}
    $MostRecentCSVTime = $OriginalCSVTime

if ($ExportFromInverter.IsPresent)
{   

    # Attempt download from inverter until retries exhausted...
    [int]$Retries = 0
    do
    {
        # Clear up any stale Sunny Explorer processes        
        get-process | where-object Name -eq "SunnyExplorer" | stop-process -force

        # Invoke Sunny Explorer to dump data for the requested date to a CSV file
        [string]$SEArgs= "`"$ExportPath\$PlantName.$PlantExtension`" -userlevel User -password $UserPwd -exportdir `"$ExportPath`" -exportrange $Date-$Date -export energy5min"
        start-process "$SEPath\SunnyExplorer.exe" $SEArgs -Wait
        start-sleep -m 1200
        
        $Retries++

        # Let's see if we've got a CSV file ...
        if (test-path "$ExportPath\$PlantName-$Date.csv")
        {
            # ... and if so, has it been updated?
            $MostRecentCSVTime = [datetime](Get-ItemProperty -Path "$ExportPath\$PlantName-$Date.csv").lastwritetime
            if ($MostRecentCSVTime -le $OriginalCSVTime)
            {
            Add-Content $UploadLog "$(Get-Date -format g) Export from Sunny Explorer for $PlantName not updated on attempt $Retries. Check Bluetooth connection to plant $PlantName."
            }
        }
        else
        {
            Add-Content $UploadLog "$(Get-Date -format g) Export from Sunny Explorer for $PlantName on $Date requested but could not be found."
        }
    }
    while (($Retries -lt 10) -and (test-path "$ExportPath\$PlantName-$Date.csv") -and ($MostRecentCSVTime -le $OriginalCSVTime))
}
else
{
    if (-not (test-path "$ExportPath\$PlantName-$Date.csv"))
        {
        Add-Content $UploadLog "$(Get-Date -format g) Export from Sunny Explorer for $PlantName on $Date not found and ExportFromInverter switch not specified."
        }
}


if ((test-path "$ExportPath\$PlantName-$Date.csv") -and (($MostRecentCSVTime -gt $OriginalCSVTime) -or -not ($ExportFromInverter.IsPresent)))
{
    # Get the contents of the CSV and create aggregate columns for the total of each inverter output (up to a maximum of 4 inverters)
    $CSVContents = Import-CSV "$ExportPath\$PlantName-$Date.csv" -delimiter ";" -Header "StatusDateTime","TotalYield0","Power0","TotalYield1","Power1","TotalYield2","Power2","TotalYield3","Power3" `
    | select-object StatusDateTime,@{Name="TotalYield";Expression={$_.TotalYield0 + $_.TotalYield1 + $_.TotalYield2 + $_.TotalYield3}},@{Name='Power';Expression={$_.Power0 + $_.Power1 + $_.Power2 + $_.Power3}}
        
    # Check the CSV looks like we're expecting
    if ($CSVContents[1].StatusDateTime -ne "Version CSV1|Tool SE|Linebreaks CR/LF|Delimiter semicolon|Decimalpoint dot|Precision 3")
    {
        Add-Content $UploadLog "$(Get-Date -format g) Unexpected information about CSV found: `"$($CSVContents[1].StatusDateTime)`". Continuing..."
    }
    
    # Check the CSV headers are the expected ones (should be line 8)
    if (($CSVContents[7].StatusDateTime -ne "dd.MM.yyyy HH:mm:ss") -or (-not ($CSVContents[7].TotalYield.StartsWith("kWh"))) -or (-not ($CSVContents[7].Power.StartsWith("kW"))))
    {
        Add-Content $UploadLog "$(Get-Date -format g) Unexpected CSV data format found: `"$($CSVContents[7].StatusDateTime);$($CSVContents[7].TotalYield);$($CSVContents[7].Power)`". Continuing..."
    }

    # Create an empty array for the statuses. The elements are the statuses unless we are in batch mode in which case the array elements are batches of statuses
    [array]$StatusArray = @()
    [int]$ArrayIndex = 0
    # For batch mode, start with an empty batch as the first element of the array
    If ($BatchMode.IsPresent)
    {
        $StatusArray = $StatusArray + ""
        [int]$BatchItemLength = 0
    }

    # Need at least 2 statuses to be able to calculate the day's yield
    if ($CSVContents.Length -ge 10)
    {
        # Work out the total historic yield at the start of the day
        [decimal]$StartYield = $CSVContents[8].TotalYield

        # Recurse through the statues from the CSV, either adding them to the array, or to the current batch
        for ([int]$StatusIndex = 9; $StatusIndex -lt $CSVContents.Length; $StatusIndex++)
        {
            # Statuses with zero incremental yields are ignored
            if ([decimal]$CSVContents[$StatusIndex].TotalYield -gt [decimal]$CSVContents[$StatusIndex-1].TotalYield)
            {
                # Calculate yield since start of day, instantaneous power, status date and time. Convert kWh and Wh to kW and W.
                [int]$Yield = ([decimal]$CSVContents[$StatusIndex].TotalYield - $StartYield) * 1000
                [int]$InstPower = ([decimal]$CSVContents[$StatusIndex].Power) * 1000
                [string]$StatusDate = Get-Date $CSVContents[$StatusIndex].StatusDateTime -format "yyyyMMdd"
                [string]$StatusTime = Get-Date $CSVContents[$StatusIndex].StatusDateTime -format "HH:mm"
                                                
                if ($BatchMode.IsPresent)
                {
                    # If insufficient room in current batch, create a new (empty) one
                    if ($BatchItemLength -eq 30)
                    {
                        $StatusArray = $StatusArray + ""
                        $BatchItemLength = 0
                        $ArrayIndex++
                    }
                    # Add status details to current batch.
                    $StatusArray[$ArrayIndex]= "$($StatusArray[$ArrayIndex])$StatusDate,$StatusTime,$Yield,$InstPower;"
                    $BatchItemLength++
                }
                else
                {
                    # Add status as a new array element
                    $StatusArray = $StatusArray + "$StatusDate,$StatusTime,$Yield,$InstPower"
                    $ArrayIndex++
                }

            }
        }
    }

    if (($BatchMode.IsPresent) -and $StatusArray[0] -ne "")
    {
        # Loop through the batches
        For ([int]$ArrayIndex = 0; $ArrayIndex -lt $StatusArray.Length; $ArrayIndex++)
        {
            # Calculate the time of the first and last statuses in the bacth
            $BatchStart = $((($StatusArray[$ArrayIndex]) -split ";")[0] -split ",")[1]
            $BatchEnd = $((($StatusArray[$ArrayIndex]) -split ";")[-2] -split ",")[1]
            try
            {
                # Post the batch of statuses to the web site
                $Response = invoke-webrequest http://pvoutput.org/service/r2/addbatchstatus.jsp -Headers @{"X-Pvoutput-SystemID"=$SystemId;"X-Pvoutput-Apikey"=$ApiKey} -Body @{"data"=$($StatusArray[$ArrayIndex])} -outfile $ResponseLog -PassThru -UseBasicParsing -UseDefaultCredentials
                Add-Content $UploadLog "$(Get-Date -format g) Successfully uploaded batch of statuses for $Date between $BatchStart and $BatchEnd - $Response"
            }
            catch
            {
                Add-Content $UploadLog "$(Get-Date -format g) Uploading batches of statuses for $Date between $BatchStart and $BatchEnd failed: $($Error[0])."
            }
            finally
            {
                # Pause before attempting next batch upload
                start-sleep 5
            }
        }
    }
    elseif ((-not ($BatchMode.IsPresent)) -and ($StatusArray.Length -ne 0))
    {
        # Loop through the statuses
        [bool]$FoundNewValues = $false
        For ([int]$ArrayIndex = 0; $ArrayIndex -lt $StatusArray.Length; $ArrayIndex++)
        {
            $StatusDate,$StatusTime,$Yield,$InstPower = $($StatusArray[$ArrayIndex] -split ",")
            # Check we haven't already uploaded these values before. (Need to use batch mode if we want to reload values.)
            if (-not (select-string -Pattern "Uploading status for $StatusDate at $StatusTime succeeded" -Path $UploadLog))
            {
                $FoundNewValues = $true
                try
                {
                    # Post the new status to the web site
                    $Response = invoke-webrequest http://pvoutput.org/service/r2/addstatus.jsp -Headers @{"X-Pvoutput-SystemID"=$SystemId;"X-Pvoutput-Apikey"=$ApiKey} -Body @{"d"=$StatusDate;"t"=$StatusTime;"v1"=$Yield;"v2"=$InstPower} -outfile $ResponseLog -PassThru -UseBasicParsing -UseDefaultCredentials
                    Add-Content $UploadLog "$(Get-Date -format g) Uploading status for $StatusDate at $StatusTime succeeded: Yield=$Yield, InstPower=$InstPower - $Response"
                }
                catch
                {
                    Add-Content $UploadLog "$(Get-Date -format g) Uploading status for $StatusDate at $StatusTime failed: Yield=$Yield, InstPower=$InstPower - $($Error[0])."
                }
                finally
                {
                    # Pause before attempting next status upload
                    start-sleep 5
                }
            }
        }
        If (-not ($FoundNewValues))
        {
            Add-Content $UploadLog "$(Get-Date -format g) Export from Sunny Explorer for $PlantName on $Date found but all statuses with positive yields have already been uploaded."
        }
    }
    else
    {
        Add-Content $UploadLog "$(Get-Date -format g) Export from Sunny Explorer for $PlantName on $Date found but contains no statuses with positive yields."
    }
}
