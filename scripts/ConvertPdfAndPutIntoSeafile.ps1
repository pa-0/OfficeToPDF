# convert current msg files into PDF

<# if($args[0] -eq $null){
	Write-Host("please specify the month number of year 2019 in the seafile storage to store the files")
	return;
}
#>
$MMdd = Get-Date -Format 'MMdd'
$MM = Get-Date -Format 'MM'
write-host "the arg is $($args[0]) and the current date is $MMdd, and will put into the month of $MM"


$childFiles = Get-ChildItem "C:\Users\wenzzha\Documents\OutlookEmailsTemp" 

foreach($file in $childFiles){
	if($file -is [System.IO.FileInfo] -and $file.Extension.Contains(".msg"))
	{
		Write-Host("processing $file to pdf")
		C:\Users\wenzzha\Documents\OutlookEmailsTemp\OfficeToPDF.exe $file  "$MMdd $($file.BaseName).pdf"
			Write-Host("converted $file to pdf")
	}
}


$newChildFiles = Get-ChildItem "C:\Users\wenzzha\Documents\OutlookEmailsTemp" 
foreach($curfile in $newChildFiles){
	if($curfile -is [System.IO.FileInfo] -and $curfile.Extension.Contains(".msg")){
		Remove-Item $curfile -Force
		Write-Host("removed $curfile")
	}
	elseif($curfile -is [System.IO.FileInfo] -and $curfile.Extension.Contains(".pdf")){
		Move-Item $curfile -Destination "C:\Users\wenzzha\Seafiles\Seafile\myfiles\2019\$MM"
		Write-Host("moved $curfile to C:\Users\wenzzha\Seafiles\Seafile\myfiles\2019\$MM")
	}
}
