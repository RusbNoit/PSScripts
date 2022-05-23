  
If ($PSVersionTable.PSVersion.Major -le 2) {
    Throw "Please execute this script from a system that has PowerShell 3.0 or newer installed."
}

# set source path with photos. Example d:\my\photos
$pathSource = 'd:\my\photos\2020'

# set slideshow directory for windows
$pathWallpaper = 'd:\WindowsSlideShow'
# 
$iMaxPhotos = 200

if (-not (test-path $pathWallpaper)) {
    Write-Host "$pathWallpaper is not exists"
    return
}

# clear old photo
Get-ChildItem "$pathWallpaper\*.jpg" | Remove-Item

# get all photo collerction
$items = Get-ChildItem $pathSource -Recurse -Include *.jpg

$itemsCount = $items.Count

if($iMaxPhotos -gt $itemsCount) {$iMaxPhotos = $itemsCount}


for ($i=0;$i -lt $iMaxPhotos;$i++) {

    $iRand = Get-Random -Maximum $itemsCount

    $sPathDest = $PSScriptRoot + "\\" + $i.ToString() + "_" + $items[$iRand].Name
    $sPathSrc = $items[$iRand].FullName

    Write-Host "Copy" $sPathSrc

    $image = New-Object -ComObject Wia.ImageFile
    $imageProcess = New-Object -ComObject Wia.ImageProcess

    [void]$image.LoadFile($sPathSrc)

    Write-Host $sPathSrc $image.Width "x" $image.Height -ForegroundColor Yellow

    if (Test-Path $sPathDest) {Remove-Item -Path $sPathDest}

    if ($image.Width -gt 300 -or $image.Height -gt 300) {

        $scale = $imageProcess.FilterInfos.Item("Scale").FilterId                    
        $imageProcess.Filters.Add($scale)
        $imageProcess.Filters[1].Properties("MaximumWidth") = 2000
        $imageProcess.Filters[1].Properties("MaximumHeight") = 2000
        $imageProcess.Filters[1].Properties("PreserveAspectRatio") = $true 

        $newimg = $imageProcess.Apply($image)
        $newimg.SaveFile($sPathDest) | Out-Null
    }

}