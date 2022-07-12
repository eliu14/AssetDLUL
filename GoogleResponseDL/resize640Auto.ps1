param ([string]$input_folder, [string] $output_folder)
function resizeIMG {
Param ( [Parameter(Mandatory=$True)] [ValidateNotNull()] $imageSource,
[Parameter(Mandatory=$True)] [ValidateNotNull()] $imageTarget,
[Parameter(Mandatory=$true)][ValidateNotNull()] $quality )

if (!(Test-Path $imageSource)){throw( "Cannot find the source image")}
if(!([System.IO.Path]::IsPathRooted($imageSource))){throw("please enter a full path for your source path")}
if(!([System.IO.Path]::IsPathRooted($imageTarget))){throw("please enter a full path for your target path")}
if ($quality -lt 0 -or $quality -gt 100){throw( "quality must be between 0 and 100.")}

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
#$Stream = New-Object System.IO.FileStream -ArgumentList $imageSource,'Open'
#$bmp  = [System.Drawing.Image]::FromStream($Stream)
$bmp = [System.Drawing.Image]::FromFile($imageSource)

#hardcoded canvas size...
$canvasWidth = 640
$canvasHeight = 640

#Encoder parameter for image quality
$myEncoder = [System.Drawing.Imaging.Encoder]::Quality
$encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
$encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($myEncoder, $quality)
# get codec
$myImageCodecInfo = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders()|where {$_.MimeType -eq 'image/jpeg'}

#compute the final ratio to use
$ratioX = $canvasWidth / $bmp.Width;
$ratioY = $canvasHeight / $bmp.Height;
$ratio = $ratioY
if($ratioX -le $ratioY){
  $ratio = $ratioX
}

#create resized bitmap
$newWidth = [int] ($bmp.Width*$ratio)
$newHeight = [int] ($bmp.Height*$ratio)
$bmpResized = New-Object System.Drawing.Bitmap($newWidth, $newHeight)
$graph = [System.Drawing.Graphics]::FromImage($bmpResized)

#$graph.Clear([System.Drawing.Color]::White)
$graph.DrawImage($bmp,0,0 , $newWidth, $newHeight)

#save to file
#$bmpResized.Save($imageTarget,$myImageCodecInfo, $($encoderParams))
$imageFmt = "System.Drawing.Imaging.ImageFormat" -as [type]
$bmpResized.Save($imageTarget.Replace('.png','l.png').Replace('.jpg','l.png').Replace('.jpeg','l.png').Replace('.JPG','l.png').Replace('.JPEG','l.png').Replace('.PNG', 'l.png'), $imageFmt::Png)

$graph.Dispose()
$bmpResized.Dispose()
$bmp.Dispose()
#$Stream.Dispose()
}

Get-ChildItem "$input_folder\*.jpg","$input_folder\*.JPG","$input_folder\*.jpeg","$input_folder\*.JPEG","$input_folder\*.png","$input_folder\*.PNG" -File | ForEach-Object {
    $newfilename = $_.Name -replace'（', '(' 
    $newfilename = $newfilename -replace '）',')'
    $newfilename = $newfilename -replace '[^i\p{IsBasicLatin}]'
    $newfilename = $newfilename -replace '[_ ]'
    echo $newfilename
    resizeIMG $_.FullName (Join-Path $output_folder $newfilename) 100
}