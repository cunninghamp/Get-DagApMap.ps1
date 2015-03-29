<#
.SYNOPSIS
Get-DagApMap.ps1 - Create a map of database copy activation preferences

.DESCRIPTION 
This PowerShell script generates a CSV report showing the activation
preferences of each of the database copies in an Exchange Server database
availability group (DAG).

.OUTPUTS
Results are output to CSV.

.EXAMPLE
.\Get-DagApMap.ps1

.NOTES
Written by: Paul Cunningham

Find me on:

* My Blog:	http://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	http://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp

For more Exchange Server tips, tricks and news
check out Exchange Server Pro.

* Website:	http://exchangeserverpro.com
* Twitter:	http://twitter.com/exchservpro

Change Log
V1.00, 30/03/2015 - Initial release
#>

#...................................
# Variables
#...................................

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path


#...................................
# Script
#...................................

#Add Exchange 2010 snapin if not already loaded in the PowerShell session
if (Test-Path $env:ExchangeInstallPath\bin\RemoteExchange.ps1)
{
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	Connect-ExchangeServer -auto -AllowClobber
}
else
{
    Write-Warning "Exchange Server management tools are not installed on this computer."
    EXIT
}


# Get the Database Availability Groups in the org

$dags = @(Get-DatabaseAvailabilityGroup)

if ($dags.count -eq 0)
{
    Write-Host "Unable to locate any database availability groups."
    EXIT
}


# Process each DAG

foreach ($dag in $dags)
{

    Write-Host -ForegroundColor White "********** Started processing DAG: $dag"


    # Initialize objects

    $report = @()
    $dagservers = @()
    $databases = @()
    $reportfilename = "$myDir\$($dag.Name)-ActivationPreferenceMap.csv"


    # Get the DAG members

    $dagservers = @($dag | Select -ExpandProperty:Servers | Sort Name)


    # Get the databases for the DAG

    $databases = @(Get-MailboxDatabase | Where {$_.MasterServerOrAvailabilityGroup -eq $dag.Name} | Sort Name)

    #Loop through the databases

    foreach ($database in $databases)
    {

        Write-Host -ForeGroundColor White "---------- Processing database: $database"

        # Get the AP values for each DB copy
        $aps = @((Get-MailboxDatabase $database).ActivationPreference)

        # Create a new object to store results
        $dbObj = New-Object PSObject -Property @{'Name'=$database.Name;}

        # Check each server for a copy of this database
        foreach ($server in $dagservers)
        {
            Write-Host "Checking $server for a copy of $database"

            if ($aps.Key.Name -icontains $server)
            {
                Write-Host -ForeGroundColor Green "Found a copy of $database on $server"

                $activationpreference = ($aps | Where {$_.Key.Name -ieq $server}).Value

                Write-Host "Activation preference is: $activationpreference"

                # Database copy found, add to custom object with AP value
                $dbObj |Add-Member NoteProperty -Name $server -Value $activationpreference
            }
            else
            {
                Write-Host "No copy of $database on $server"
                
                # No database copy found, add to custom object with blank AP value
                $dbObj |Add-Member NoteProperty -Name $server -Value ""
            }

        }

        # Add the custom object to the report
        $report += $dbObj

    }

    # Output the report for this DAG to CSV
    $report | Sort Name | Export-CSV -NoTypeInformation -Path $reportfilename -Encoding UTF8

    Write-Host -ForegroundColor White "********** Finished processing DAG: $dag"
}

#...................................
# Finishsed
#...................................