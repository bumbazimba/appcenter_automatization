#This script builds several branches for an Application in AppCenter and saves a report.
#Check Readme.md file for more details - https://github.com/bumbazimba/appcenter_automatization/blob/master/README.md
#Author Nikita Nikolaev
#Version 1.0

#Enter AppCenter credentials for API requests
	$user = Read-Host "Enter AppCenter username"
	$app = Read-Host "Enter Application name"
	$token = Read-Host "Enter API Token"
	[int]$branches_number = Read-Host "Enter number of branches to build"
    $branches_number = $branches_number-1

#Create Excel file and insert table headers
	$excel = New-Object -ComObject excel.application
	$excel.visible = $True
	$workbook = $excel.Workbooks.Add()
	$sheet= $workbook.Worksheets.Item(1)
	$sheet.Name = 'Report'
	$sheet.Cells.Item(1,1)= 'Branch name'
	$sheet.Cells.Item(1,2)= 'Build status'
	$sheet.Cells.Item(1,3)= 'Duration (min)'
	$sheet.Cells.Item(1,4)= 'Link to build logs'
	$row = 2

#Get All branches from the Application and save branches names in array
	$branches = Invoke-RestMethod -Uri "https://api.appcenter.ms/v0.1/apps/$user/$app/branches" -Method Get `
		-Headers @{"Accept"="application/json"; "X-API-Token"="$($token)"; "Content-Type"="application/json"};
	Write-Host "`nList of all branches in Application $app :"
	$branches.branch.name
	$branches = $branches.branch.name

#Build First $branches_number branches and save IDs for next step
	$build_id = @()

	Foreach ($branch in $branches[0..$branches_number])	
		{ 		
			$build_id += Invoke-RestMethod -Uri "https://api.appcenter.ms/v0.1/apps/$user/$app/branches/$branch/builds" -Method Post `
				-Headers @{"Accept"="application/json"; "X-API-Token"="$($token)"; "Content-Type"="application/json"} | Select-object id
			Write-host "Branch $branch - build started"
			$sheet.Cells.Item($row,1)=$branch
			$row++
		}
#Get Duration, Status and Log file link for every build by ID
	$row=2
	$result = "in progress"
	$number = 0
	Write-host "`n"
		Foreach ($id in $build_id)
		{
			$id = $build_id[$number].id
			#.API request of build's details. If build still in progress, wait 30 seconds
				while (($result -ne "failed") -and ($result -ne "success"))
				{
					$details = Invoke-RestMethod -Uri "https://api.appcenter.ms/v0.1/apps/$user/$app/builds/$id" -Method Get `
						-Headers @{"Accept"="application/json"; "X-API-Token"="$($token)"; "Content-Type"="application/json"};
					$result = $details.result
                    Write-Host "Build ID $id is in progress"
					Start-Sleep -s 30
				}
				#Calculating Duration time
					$starttime = $details.startTime
					$finishtime = $details.finishTime
					$diff= New-TimeSpan -Start $starttime -End $finishtime
					$duration = $diff.ToString()
				#Generate log download link	
					$log = "https://appcenter.ms/download?url=%2Fv0.1%2Fapps%2F"+$user+"%2F"+$app+"%2Fbuilds%2F"+$id+"%2Fdownloads%2Flogs"
				#Insert data to Excel table
					$sheet.Cells.Item($row,2)=$result
					$sheet.Cells.Item($row,3)=$duration
					$sheet.Cells.Item($row,4)=$log
					$row++
					$number++
					$result = "in progress"
					Write-Host "Build ID $id completed!"
		}

#Resize Excel columns and save report on Desktop
	Write-Host "`nAll builds finished, please check status in the Report" -ForegroundColor Green
	Write-Host "`nReport was saved on your Desktop in Excel file!" 
	$date = Get-Date -Format "MM/dd/yyyy_HH_mm_ss"
	$output_path = "C:\Users\$env:username\Desktop\Report_$date.xlsx"
	$usedRange = $sheet.UsedRange
	$usedRange.EntireColumn.AutoFit() | Out-Null
	$workbook.SaveAs($output_path)
	$excel.Quit()


