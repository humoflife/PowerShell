<#
.SYNOPSIS
	Displays the largest files.
.DESCRIPTION
	Get-LargestFiles finds and displays the largest files. By default the number of files displayed is 10 and the default start directory is $HOME. Both parameters can be changed.
.PARAMETER Dir
	Whether to obtain the list of computers from Active Directory. It requires the systems to be member of an Active Directory domain. 
.PARAMETER NumFiles
    The list of computers. If a computer file exists with computer names, the command line list is added to the list provided in the computers text file.
.EXAMPLE
    C:\PS>Get-LargestFiles.ps1 -NumFiles 20
	Displays the largest 20 files in the $HOME directory
.NOTES
    Author: Markus Schweig
    Date:   November 5, 2011
.LICENSE
	This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

    This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
#>

param (
	[string] $Dir = $HOME,
	[int] $NumFiles = 10
)

Dir $Dir -Recurse -ea SilentlyContinue | Sort-Object Length -Descending | Select-Object -First $NumFiles | Select-Object FullName, Length, LastWriteTime