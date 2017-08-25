cls
$cwd = (Resolve-Path .\).Path

Write-host "Downloading latest Patching information...." -fo red
$downloadLink = (((Invoke-RestMethod -uri https://www.microsoft.com/en-gb/download/confirmation.aspx?id=36982).ToString() -split ">" | select-string "BulletinSearch.xls" | Select-String " href")[0] -split '"')[1]
Invoke-WebRequest -Uri $downloadLink -OutFile "BulletinSearch.xls"

Function ExcelCSV ($File)
{
 
    $excelFile = "$pwd\" + $File + ".xls"
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $wb = $Excel.Workbooks.Open($excelFile)
    foreach ($ws in $wb.Worksheets)
    {
        $ws.SaveAs("$pwd\" + $File + ".csv", 6)
    }
    $Excel.Quit()
}
$FileName = "BulletinSearch"

Write-host "Converting Patching information....." -fo Green

ExcelCSV -File "$FileName"
Remove-Item *.xls

Write-host "Enter Selection" -fo red
"
1. Collate list of all OS Patches
2. Lookup KB Number (KBxxxxxxx)
3. Check Bullentin Number (MSxx-xxx)
"
do {
$choice = Read-Host "
Query Options: (1-3)"
} until ($Choice -in ("1","2","3","4"))
if ($choice -eq "1") {$Quest = "Affected_Product"}
if ($choice -eq "2") {$Quest = "Bulletin_KB"}
if ($choice -eq "3") {$Quest = "Bulletin_Id"}
$choice = $null


$KBPrompt = Read-Host "Enter Search Term" 
$SupMis = Import-Csv $cwd\$filename.csv -header Date_Posted,Bulletin_Id,Bulletin_KB,Severity,Impact,Title,Affected_Product,Component_KB,Affected_Component,Impact2,Severity2,Supersedes | Where-Object {$_.$Quest -like '*'+$KBPrompt +'*'} 

$SupMis 

Write-host "Compiling a list of Superseded Patches....." -fo Green
Write-host "==========================================================================" -fo Blue
Write-host "The following is a list of patches that have been superseded by " -fo Green
Write-host "                           "$KBPrompt -fo red
Write-host "==========================================================================" -fo Blue

$OutFile = $SupMis | Where-Object {$_.$Quest -like '*'+$KBPrompt +'*'} | Select-Object Supersedes | Sort-Object supersedes -Unique | Format-List -Property *
$outfile
$Note = "A List of superseded patches based on search term : " + $KBPrompt | Out-File Superseded.txt
$OutFile | Out-File Superseded.txt -Append
Pause
