<#
.SYNOPSIS
	Displays event count information for requested computers.
.DESCRIPTION
	Get-AsynchEventCount offers multiple ways to specify the computer names for which event count information shall be collected, including a file with computer names, a command line list of computer names, and an automated lookup through Active Directory for those environments where computers reside in a domain. It performs a computer name to address lookup and excludes those computer names from the requested list that cannot be resolved. For the final list of computers, it establishes sessions on these computers to individually collect their application and system event log count information. If the sessions cannot be established which is possible due to security restrictions, the script informs the administrator to verify secuirty settings and it exits. Otherwise, it returns a table with the requested event count information for each computer. The purpose of the script is to allow IT and LiveOps groups to focus on driving chronic issues out of their environment by resolving those issues first that occur the most across all inspected computers across the entire environment and thereafter those that occur at high values on specific computers. By default, the script obtains event count information for the past 24 hours. The timeframe can be change through the -After and -Before parameters.
	
	Note: The computers on which the event count scriptblock runs need to be configured to trust the computer from which the script is run. On Windows 7 and Windows Server 2008 systems it should suffice to include the name of the computer that runs the script in the TrustedHost winrm client configuration through winrm s winrm/config/client '@{TrustedHost="YourTrustedHost"}' and then running a quick configuration with winrm quickconfig. Verify that the winrm service is running on the remote computers through get-service -ComputerName YourComputerName winrm. Set-executionPolicy to a desired value for a non-signed copy of this script or sign it with Set-AuthenticodeSignature and your code signing certificate.
.PARAMETER AD
	Whether to obtain the list of computers from Active Directory. It requires the systems to be member of an Active Directory domain. 
.PARAMETER ComputerList
    The list of computers. If a computer file exists with computer names, the command line list is added to the list provided in the computers text file.
.PARAMETER ComputerFile
    A text file that contains the list of computers. The script also looks into the computers.txt default file inside the same directory where the script resides if no -ComputerFile parameter is given. If the file does not exist, its absence is silently ignored.
.PARAMETER Before
	The time up until events should be counted
.PARAMETER After
	The time after which events should be counted
.PARAMETER EntryType
	The event log entry type for the events that should be counted. The default is "error". Other values can include "warning" or "information".
.PARAMETER Rows
	A postive value that limits the number of rows that will be displayed to only focus on the highest event counts. If the value is greater than the number of distinct events for which there is a count then all rows are displayed. A value of 0 will display all rows.
.PARAMETER CSV
	Save the output to the "eventCounts.csv" file
.PARAMETER Silent
	Suppresses all output except for the final event count table. It also silences -Verbose output.
.PARAMETER Verbose
	Displays extra chatty output. -Silence takes precedence over -Verbose when both parameters are used.
.EXAMPLE
    C:\PS>Get-AsynchEventCount
    This produces the event count for the local system for the last 24 hours.
.EXAMPLE
    C:\PS>Get-AsynchEventCount -AD
    This produces the event count for all systems obtained from Active Directory for the current domain for the last 24 hours.
.EXAMPLE
    C:\PS>Get-AsynchEventCount -ComputerList c1, c2 -After "10/20/2011 17:00"
    This produces the event count for the systems named c1 and c2 since 10/20/2011 17:00 up to the time when this script runs.
.EXAMPLE
    C:\PS>Get-AsynchEventCount -AD -After "10/20/2011 17:00" -Rows 3 -EntryType warning -CSV
    This produces the top 3 rows of warning event counts for the systems obtained through Active Directory since 10/20/2011 17:00 up to the time when this script runs and writes the output to the command line window and "eventCounts.csv".
.NOTES
    Author: Markus Schweig
    Date:   November 1, 2011
.LICENSE
	This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

    This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
#>

param (
	[switch] $AD,
	[string] $ComputerFile = "computers.txt",
	[string[]] $ComputerList = $null,
	[string] $Before = $(Get-Date),
	[string] $After = $([DateTime]::Now.AddHours(-24)),
	[string] $EntryType = "error",
	[int] $Rows = 0,
	[switch] $CSV,
	[switch] $Silent,
	[switch] $Verbose
)

$requestedComputers = @()
$computers = @()

# validate $EntryType
if ("Error", "Warning", "Information" -NotContains $EntryType) {
	if (!$Silent) {
		Write-Host "Please specify a valid -EntryType of Error, Warning, or Information."
	}
	exit(-1)
}

# get computer names from a specified file, or the default file "computers.txt"
if ($ComputerFile) {
	if ((Test-Path $ComputerFile)) {
		if ($Verbose -and !$Silent) {
			Write-Host "Getting computer names from"$ComputerFile"."
		}
		$requestedComputers = Get-Content $ComputerFile| where-object {$_ -notmatch "#.+$"}
	} else {
		if ($ComputerFile -ne "computers.txt") {
			if (!$Silent) {
				Write-Host "Please specify a valid ComputerFile. "$ComputerFile" does not exist."
			}
			exit(-1)
		} else {
			$ComputerFile = $null
		}
	}
}

# add computer names from the command line
if ($ComputerList) {
	foreach ($computer in $ComputerList) {
		if ($requestedComputers -notcontains $computer) {
			$requestedComputers += $computer
		}
	}
}

# try to obtain Active Directory computer information
if ($AD) {
	if ($Verbose -and !$Silent) {
		Write-Host "Obtaining computers from Active Directory"
	}
	try {
		$ADcomputers = Get-ADcomputer -Filter *|Format-Table -Property Name -HideTableHeaders|Out-String -Stream
		foreach ($computer in $ADcomputers) {
			if ($computer.Length -eq 0) {
				continue
			}
			$requestedComputers += $computer.Trim()
		}
	} catch {
		if (!$ComputerFile -and !$ComputerList) {
			if (!$Silent) {
				Write-Warning "Unable to obtain Active Directory computer list. Is your computer part of a domain?"
			}
			exit (-1)
		} else {
			if (!$Silent) {
				Write-Warning "Unable to obtain Active Directory computer list. Processing the non-AD information."
			}
		}
	}
}

# if there is no list of computer names through any of the available options, we use the local computer as default
if (!$ComputerList -and !$ComputerFile -and !$AD) {
	$requestedComputers = $env:COMPUTERNAME
}

# validate requested computers name resolution and only keep those that resolve
foreach ($computer in $requestedComputers) {
	try {
		$result = [System.Net.Dns]::GetHostByName($computer)
		$computers += $computer
	} catch {
		if (!$Silent) {
			Write-Host Removing $computer from requested list because its name cannot be resolved.
		}
	}
}

# if after validation the computer list is empty, we exit
if ($computers.Length -eq 0) {
	if (!$Silent) {
		Write-Warning "There are no resolvable computer names."
	}
	exit(-1)
}

if ($Verbose -and !$Silent) {
	Write-Host "Testing output for" $computers"."
}

if ($Verbose -and !$Silent) {
	Write-Host "Creating remote sessions."
}
$results = @()
$s1 = new-pssession -computername $computers
$s2 = new-pssession -computername $computers

if ($s1 -eq $null -or $s2 -eq $null) {
	if (!$Silent) {
		Write-Host "Unable to establish remote sessions. If not yet, try running this script with elevated permissions and check winrm settings on" $computers"."
	}
	exit (-1)
}

$null = invoke-command -session $s1 -scriptblock {param($InnerAfter, $InnerBefore, $InnerEntryType) & get-eventlog -LogName system -entrytype $InnerEntryType -after $InnerAfter -before $InnerBefore|Group-Object -Property EventId -NoElement
} -ArgumentList $After, $Before, $EntryType -AsJob
$null = invoke-command -session $s2 -scriptblock {param($InnerAfter, $InnerBefore, $InnerEntryType) & get-eventlog -LogName application -entrytype $InnerEntryType -after $InnerAfter -before $InnerBefore|Group-Object -Property EventId -NoElement
} -ArgumentList $After, $Before, $EntryType -AsJob
$jobs = Get-Job
$null = Wait-Job -Job $jobs
if ($Verbose -and !$Silent) {
	Write-Host "Remote jobs completed."
}
$results = Receive-Job -ComputerName $computers -Job $jobs
$eventIds = $results|Sort-Object -Property Name -Unique|Select-Object -Property Name

try {
	Remove-Job -Job $jobs
} catch {
	# Sometimes the runtime system has removed the jobs already
}
Remove-PSSession -Session $s1
Remove-PSSession -Session $s2

# build table presentation object
if ($Verbose -and !$Silent) {
	Write-Host "Processing event count data for tabular presentation."
}
$eventTable = New-Object system.Data.DataTable “EventTable”
$eventIdCol = New-Object system.Data.DataColumn EventId,([Int32])
$eventTable.columns.add($eventIdCol)
foreach ($computer in $computers) {
	$col = New-Object system.Data.DataColumn $computer,([int])
	$eventTable.columns.add($col)
}
$eventIdTotalCol = New-Object system.Data.DataColumn Total,([Int32])
$eventTable.columns.add($eventIdTotalCol)

# fill the table cells with the eventid counts
foreach ($eventid in $eventIds) {
	$total = 0
	$row = $eventTable.NewRow()
	$row.EventId = [Int32]$eventId.Name
	foreach ($computer in $computers) {
		$row.$computer = [Int32]($results|Where-Object {$_.Name -eq $eventId.Name -and $_.PSComputerName -eq $computer}).Count
		$total += $row.$computer
	}
	$row.Total = $total
	$eventTable.Rows.Add($row)
}

if (!$Silent) {
	Write-Host
	Write-Host Event Count from $after to $before
}

# remove table rows from the bottom if desired
if ($Rows -eq 0) {
	$Rows = $eventTable.Rows.Count
}
$eventTable|Sort-Object -Property Total -Descending|Select-Object -First $Rows|Format-Table -AutoSize
if ($CSV) {
	$eventTable|Sort-Object -Property Total -Descending|Select-Object -First $Rows|ConvertTo-Csv -NoTypeInformation|Out-File "eventCounts.csv"
}