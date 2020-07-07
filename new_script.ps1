#This script builds several branches for an Application in AppCenter and saves a report.
#Check Readme.md file for more details - https://github.com/bumbazimba/appcenter_automatization/blob/master/README.md
#Designed for Powershell 6 or above. Tested on Powershell 7
#Author Nikita Nikolaev
#Version 1.1

#Enter AppCenter credentials for API requests

[string]$user = Read-Host "Enter AppCenter username"
[string]$app = Read-Host "Enter Application name"
[string]$token = Read-Host "Enter API Token"
[int]$branches_number = Read-Host "Enter number of branches to build"
$branches_number = $branches_number-1


$excel = Create-ExcelFile
$branches = Get-Branches
$build = Build-Branches $branches
$build_details = Get-BuildDetails $build
Insert-ExcelFile $excel, $build_details, $branches
Save-ExcelFile $excel



#Get All branches from the Application
function Get-Branches ()
{
	$branches = Invoke-RestMethod -Uri "https://api.appcenter.ms/v0.1/apps/$user/$app/branches" -Method Get `
		-Headers @{"Accept"="application/json"; "X-API-Token"="$($token)"; "Content-Type"="application/json"} -MaximumRetryCount 5 -RetryIntervalSec 2;
	Write-Host "`nList of all branches in Application $app :"
    Write-Host $branches
    return $branches
}


#Build specified number of branches
function Build-Branches ()
{
    param
    (
        [Parameter(Mandatory=$true)][object]$branches
    )

	$build = @()
    
	#Check if user wants to build All or just few branches
	[int]$branches_count = $branches.Length - 1
    $branches_names = $branches.branch.name
        if ($branches_count -ne $branches_number)
        {
            $branches_count = $branches_number
        }
        else 
        {
            $branches_count = $branches_count - 1
            $branches_number = $branches_count
        }
	Foreach ($branch in $branches_names[0..$branches_count])	
		{ 		
			$build += Invoke-RestMethod -Uri "https://api.appcenter.ms/v0.1/apps/$user/$app/branches/$branch/builds" -Method Post `
				-Headers @{"Accept"="application/json"; "X-API-Token"="$($token)"; "Content-Type"="application/json"} `
				-MaximumRetryCount 5 -RetryIntervalSec 2 | Select-object id
			Write-Host "Branch $branch - build started"
		}
        return $build
}

#Get Details about all Builds
function Get-BuildDetails ()
{
        param
    (
        [Parameter(Mandatory=$true)][object]$build
    )
	
	$result = "in progress"
	$number = 0
    $build_result = @()
		Foreach ($id in $build)
		{
			$id = $build[$number].id
			#API request of build's details. If build is still in progress, wait 20 seconds
				while (($result -ne "failed") -and ($result -ne "success"))
				{
					$details = Invoke-RestMethod -Uri "https://api.appcenter.ms/v0.1/apps/$user/$app/builds/$id" -Method Get `
						-Headers @{"Accept"="application/json"; "X-API-Token"="$($token)"; "Content-Type"="application/json"} `
						-MaximumRetryCount 5 -RetryIntervalSec 2;
					$result = $details.result
                    Write-Host "Build ID $id is in progress"
					Start-Sleep -s 20
				}
                $build_result += $details
				$number++
				$result = "in progress"
				Write-Host "Build ID $id completed!"            
		}
        return $build_result
}


#Create Excel file and insert table headers
function Create-ExcelFile ()
{

	$excel = New-Object -ComObject excel.application
	$excel.visible = $True
    $workbook = $excel.Workbooks.Add()
	$sheet= $workbook.Worksheets.Item(1)
	$sheet.Name = 'Report'
	$excel.Cells.Item(1,1)= 'Branch name'
	$excel.Cells.Item(1,2)= 'Build status'
	$excel.Cells.Item(1,3)= 'Duration (min)'
	$excel.Cells.Item(1,4)= 'Link to build logs'
    return $excel
}

#Insert build details in Excel file
function Insert-ExcelFile ()
{  
        param
    (
        [Parameter(Mandatory=$true)][object]$excel,
        [int]$row = 2
    )  

    #Get and Insert Branch names in Excel file
        $name = $branches.branch.name     
        Foreach ($id in $name[0..$branches_number])
        {
            $excel.ActiveWorkbook.ActiveSheet.Cells.Item($row,1) = $id
            $row++
        }

    #Get and Insert other build Details in Excel file
        $row = 2
        Foreach ($build in $build_details) 
        {    
            $result=$build.result
        #Calculate Duration time
	        $starttime = $build.startTime
	        $finishtime = $build.finishTime
	        $diff= New-TimeSpan -Start $starttime -End $finishtime
	        $duration = $diff.ToString()

	    #Generate log download link	
            $build_id = $build.id
	        $log = "https://appcenter.ms/download?url=%2Fv0.1%2Fapps%2F"+$user+"%2F"+$app+"%2Fbuilds%2F"+$build_id+"%2Fdownloads%2Flogs"

	    #Insert data to the Excel table
	        $excel.ActiveWorkbook.ActiveSheet.Cells.Item($row,2)=$result
	        $excel.ActiveWorkbook.ActiveSheet.Cells.Item($row,3)=$duration
	        $excel.ActiveWorkbook.ActiveSheet.Cells.Item($row,4)=$log
	        $row++
        } 
}

#Resize Excel columns and save report on Desktop
function Save-ExcelFile ()
{
    param
    (
        [object]$excel
    )
    Write-Host "`nAll builds finished, please check status in the Report"
	Write-Host "`nReport was saved on your Desktop in Excel file!" 
	$date = Get-Date -Format "MM/dd/yyyy_HH_mm_ss"
	$excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
	$excel.ActiveWorkbook.SaveAs("C:\Users\$env:username\Desktop\Report_$date.xlsx")
	$excel.Quit()
}


