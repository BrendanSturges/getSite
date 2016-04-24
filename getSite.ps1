<#

Script created by Brendan Sturges, reach out if you have any issues.
This script queries a file the user chooses and checks all servers within to see which sccm site this server is connected to and outputs that info to a file that the user specifies.

#>


Function Get-FileName($initialDirectory){   
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
	Out-Null

	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.initialDirectory = $initialDirectory
	$OpenFileDialog.filter = "All files (*.*)| *.*"
	$OpenFileDialog.ShowDialog() | Out-Null
	$OpenFileDialog.filename
}

Function Save-File([string] $initialDirectory ) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() |  Out-Null
	
	$nameWithExtension = "$($OpenFileDialog.filename).csv"
	return $nameWithExtension

}

#Open a file dialog window to get the source file
$serverList = Get-Content -Path (Get-FileName)

#open a file dialog window to save the output
$fileName = Save-File $fileName


#define "i" for progress bar
$i = 0
$ErrorActionPreference = 'Stop'

foreach($server in $serverList)
{

	$ErrorMessage = ''
	$siteCode = ''

	Try {
		$siteCode = (Invoke-WMIMethod -computername $Server -namespace root\ccm -Class SMS_Client -Name GetAssignedSite).sSiteCode
			$props = [ordered]@{
			'Server' = $server
			'Site' = $siteCode
			'Details' = ''
		}
		$obj = New-Object -TypeName PSObject -Property $props

	}
	
	Catch {
		if(Test-Connection -ComputerName $server -Count 2 -Quiet)
			{
			$ErrorMessage = $_.Exception.Message
			}
		else
			{
			$ErrorMessage = 'Server is Offline'
			}
		
		$props = [ordered]@{
			'Server' = $server
			'Site' = 'ERROR'
			'Details' = $ErrorMessage
		}
			
		$obj = New-Object -TypeName PSObject -Property $props
	
	}
	
	Finally {
		$data = @()
		$data += $obj
		$data | Export-Csv $fileName -noTypeInformation -append	
	}
	$i++
	Write-Progress -activity "Checking server $i of $($serverList.count)" -percentComplete ($i / $serverList.Count*100)	
		
}

