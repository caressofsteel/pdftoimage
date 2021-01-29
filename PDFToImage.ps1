###---------------------------------------------------------------
### Filename:           PDFToImage.ps1
### Author:             David Rodgers
### Date:               2020.01.27
###---------------------------------------------------------------
### Extract and convert pages from a PDF to various image formats.
###
### Uses the following command line tools:
### - Exiftool
### - Poppler (pdftocairo
###---------------------------------------------------------------

### Extract command output of both STDERR/STDOUT
### https://jackgruber.github.io/2018-05-11-ps-get-process-output/
function Get-ProcessOutput
{
    Param (
                [Parameter(Mandatory=$true)]
                $FileName,
                $Args
    )
    
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo.UseShellExecute = $false
    $process.StartInfo.RedirectStandardOutput = $true
    $process.StartInfo.RedirectStandardError = $true
    $process.StartInfo.FileName = $FileName
    if($Args) { $process.StartInfo.Arguments = $Args }
    $out = $process.Start()
    
    $StandardError = $process.StandardError.ReadToEnd()
    $StandardOutput = $process.StandardOutput.ReadToEnd()
    
    $output = New-Object PSObject
    $output | Add-Member -type NoteProperty -name StandardOutput -Value $StandardOutput
    $output | Add-Member -type NoteProperty -name StandardError -Value $StandardError
    return $output
}

# Create output folder
$OutputPath = "$PSScriptRoot\output\"
If(!(test-path $OutputPath))
{
    New-Item -ItemType Directory -Force -Path $OutputPath
}

# Create source folder
$SourceFolder = "$PSScriptRoot\pdf\"
If(!(test-path $SourceFolder))
{
      New-Item -ItemType Directory -Force -Path $SourceFolder
}

# Set alternate source folder for PDF files
# $SourceFolder = "$PSScriptRoot\test\"

# Get a list of the PDF files
$files = Get-ChildItem "$SourceFolder" -Filter *.pdf
$totalfiles = $files.Length

# Create report object
$OutArray = @()

# Update console
Clear-Host
Write-Host "---------------------------------------------------"
Write-Host PDFToImage - Processing $totalfiles PDFs
Write-Host "---------------------------------------------------"

# Process each PDF in source folder
foreach ($f in $files) {
    
    $basefilename = [IO.Path]::GetFileNameWithoutExtension( $f )
    $Arg = '-PageCount "' + $SourceFolder + $f + '"'
    $PDFIndex = $files.indexof($f)
    
    # Rename files for easy indexing
    if ($PDFIndex -le 8) {
        $CurrentPDF = "0" + ($PDFIndex + 1)
    } else {
        $CurrentPDF = $PDFIndex + 1
    }

    # Count pages in PDF using Exiftool
    $CountPages = Get-ProcessOutput -FileName ".\exiftool.exe" -Args $Arg
    
    # Extract page count from Exiftool output
    $TotalPages = (-split $CountPages).ForEach({ "$_" })   

    # Begin PDF processing
    Write-Host Processing: [$CurrentPDF] $f with $TotalPages[3] pages
    
    # Extract first page (Cover)
    .\pdftocairo -f 1 -l 1 -jpeg -q $SourceFolder/$f $OutputPath/[$CurrentPDF]-$basefilename-01-Cover
    
    # Extract last page (Back)
    .\pdftocairo -f $TotalPages[3] -l $TotalPages[3] -jpeg -q $SourceFolder/$f $OutputPath/[$CurrentPDF]-$basefilename-02-Back
    
    # Add to report object
    $report = "" | Select-Object "Number", "Title", "Pages"

    # Fill the object
    $report.number = $CurrentPDF
    $report.title = $f
    $report.pages = $TotalPages[3]

    # Add the object to the out-array
    $outarray += $report

    # Wipe the object just to be sure
    $report = $null    
}

# After the loop, export the array to CSV and backup the old one
Write-Host "---------------------------------------------------"
Write-Host Finished processing ($PDFIndex + 1) files
Write-Host Creating Report
$fileToCheck = "$PSScriptRoot\output\pdftoimage.csv"
if (Test-Path $fileToCheck -PathType leaf) 
{
    Remove-Item "$OutputPath\pdftoimage.csv"
    $outarray | export-csv "$OutputPath\pdftoimage.csv" -NoTypeInformation
}
else
{
    $outarray | export-csv "$OutputPath\pdftoimage.csv" -NoTypeInformation
}
Write-Host Done!
pause

###---------------------------------------------------------------
### END OF LINE
###---------------------------------------------------------------  