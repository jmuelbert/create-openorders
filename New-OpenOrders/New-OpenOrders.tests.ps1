$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path).Replace(".Tests.", ".")
. "$here${directorySeparatorChar}$sut"

function Get-Tempdir() {
    $TempName = New-TemporaryFile
    $outDir = Split-Path -Path $TempName
    $fileName = Split-Path -Leaf $TempName
    $dirName = $fileName.Split(".")
    $outPath = $outDir + "\" + $dirName[0]
    New-Item $outPath -type Directory
    Write-Log "The XLSX where generated in : $OutPath"
    return $outPath
}

Describe "New-OpenOrders" {
    It "New-OpenOrders-full" {
        New-OpenOrders -usersfile berater.csv -datafile Data.xslx -outputPath Get-Tempdir
    }
}
