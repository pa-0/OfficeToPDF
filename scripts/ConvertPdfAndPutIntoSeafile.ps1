# convert current msg files into PDF

<# if($args[0] -eq $null){
	Write-Host("please specify the month number of year 2019 in the seafile storage to store the files")
	return;
}
#>
$MMdd = Get-Date -Format 'MMdd'
$MM = Get-Date -Format 'MM'
write-host "the arg is $($args[0]) and the current date is $MMdd, and will put into the month of $MM"

$RootPath = "D:\OutlookEmailsTemp"
$SeafileRootPath = "D:\Seafile\Seafile\myfiles"
$childFiles = Get-ChildItem $RootPath

foreach($file in $childFiles){
	if($file -is [System.IO.FileInfo] -and $file.Extension.Contains(".msg"))
	{
		Write-Host("processing $file to pdf")
		& $RootPath"\OfficeToPDF.exe" $file  "$MMdd $($file.BaseName).pdf"
			Write-Host("converted $file to pdf")
	}
}

$newChildFiles = Get-ChildItem $RootPath
foreach($curfile in $newChildFiles){
	if($curfile -is [System.IO.FileInfo] -and $curfile.Extension.Contains(".msg")){
		Remove-Item $curfile -Force
		Write-Host("removed $curfile")
	}
	elseif($curfile -is [System.IO.FileInfo] -and $curfile.Extension.Contains(".pdf")){
		Try{
			Move-Item $curfile -Destination "$SeafileRootPath\2019\$MM" -ErrorAction Continue
		}Catch
		{
			#usually it is caused by the long file name.
			$length = $curfile.BaseName.Length
			$newfileName = "$($curfile.BaseName.substring(0, 50))---$($curfile.BaseName.substring($length-50)).pdf"
			Rename-Item $curfile -NewName $newfileName
			Move-Item $newfileName -Destination  "$SeafileRootPath\2019\$MM"
		}
		Write-Host("moved $curfile to $SeafileRootPath\2019\$MM)")
	}
}