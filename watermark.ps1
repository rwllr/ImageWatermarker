# Created by Tony Gustafsson
# Version 1.0
# Release date 2012-05-03

#--------------User defined varables-------------------
$imagesPath = "C:\bilder\Produkter\*";											#The path of images to watermark
$extensions = "*.jpg";															#Which filetypes to watermark
$newerThan = 7;																	#Number of days, older than this the images will be ignored to save time
$exifIdentifier = "WaterMarked";												#The value to identify watermarked images in EXIF
$watermarkImage = "C:\ThisScriptLocation\watermark.png";						#The watermark image to add to the images
$logFile = "C:\ThisScriptLocation\logs\ " + $(Get-Date -format YY-MM-DD-hhmmss) + ".txt"		#Error log file
$tmpFile = "C:\tmp_watermarked";												#Filename and dir to temporary file
#------------------------------------------------------

#Load assemblys from .NET
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing");
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Text");

Function Write-ImageWaterMark {
	#The function for adding watermarks to the images
	Param (
		[Parameter(ValueFromPipeline=$False, Mandatory=$True, HelpMessage="A path to original image")]
		[string]$SourceImage,

		[Parameter(ValueFromPipeline=$False, Mandatory=$True, HelpMessage="A path to watermark image")]
		[string]$watermarkImage,

		[Parameter(ValueFromPipeline=$False, Mandatory=$True, HelpMessage="A path to target image")]
		[string]$TargetImage
	);
	
	#Exception handeling, log errors to file
	trap {
		$(Get-Date -format T) + " " + $_.Exception.Message >> $logFile;
		continue;
	}
	
	#Load image from file
	$theImage = [System.Drawing.Image]::FromFile($SourceImage);

	#Load watermark from file
	$watermarkImg = [System.Drawing.Image]::FromFile($watermarkImage);

	#Create a graphics area from on the real image
	$imageGraphics = [System.Drawing.Graphics]::FromImage($theImage);
	
	#Create a paint brush with the watermark on it
	$watermarkBrush = New-Object System.Drawing.TextureBrush($watermarkImg);
	$watermarkBrush.WrapMode = [System.Drawing.Drawing2D.WrapMode]::Clamp; #Tile Or Clamp
	$newMatrix = New-Object System.Drawing.Drawing2D.Matrix;
	$newMatrix.Translate(($theImage.width - $watermarkImg.width), ($theImage.height - $watermarkImg.height), [System.Drawing.Drawing2D.MatrixOrder]::Append);
	$watermarkBrush.Transform = $newMatrix;
	
	#Add the paint brush to the graphics
	$newRect = New-Object System.Drawing.Rectangle(0, 0, $theImage.Width, $theImage.Height);
	$imageGraphics.FillRectangle($watermarkBrush, $newRect);
	
	#Adding EXIF data for identification
	$property = $theImage.PropertyItems[0];
	$property.Id = 0x9213;
	$property.Type = 2;
	$property.Value = [System.Text.Encoding]::UTF8.GetBytes($exifIdentifier);
	$property.Len = $property.Value.Count;
	$theImage.SetPropertyItem($property);
	
	#Save the image to a temporary file, cannot replace the original file
	$theImage.Save($tmpFile, [System.Drawing.Imaging.ImageFormat]::Jpeg);
	
	#Release the files, to this point the files is locked
	$theImage.Dispose();
	$imageGraphics.Dispose();
	$watermarkImg.Dispose();
	 
	#Garbage collector
	[gc]::collect();
	[gc]::WaitForPendingFinalizers();

	#Replace the original file with the temporary one
	Remove-Item $TargetImage;
	Move-Item $tmpFile $TargetImage;	
}

Function Read-ImageWaterMark([string]$SourceImage) {
	#For reading EXIF property "ImageHistory", which is set to a value to indicate
	#that the image has been watermarked before.
	#Also resaving as JPEG if the images doesn't have 24 bit colors
	$image = New-Object System.Drawing.Bitmap $SourceImage;

	try {
		$image_history_prop = $image.GetPropertyItem(0x9213);
		$image_history = [System.Text.Encoding]::ASCII.GetString($image_history_prop.Value);
	}
	catch {
		$image_history = "notWatermarked"; #Or whatever exept $exifIdentifier
	}
	
	if ($image.PixelFormat.ToString().Contains("Indexed")) {
		#If the images doesn't have 24 bit colors, resave as real JPEG
		#so that Write-ImageWaterMark won't get any problems.
		Write-Host "Error: $SourceImage did not have 24 bit colors, trying to save it in the right format.";
		$(Get-Date -format T) + " Error: $SourceImage did not have 24 bit colors, trying to save it in the right format." >> $logFile;
		$image.Save($tmpFile, [System.Drawing.Imaging.ImageFormat]::Jpeg);
		$image.Dispose();

		Remove-Item $SourceImage;
		Move-Item $tmpFile $SourceImage;
	}
	
	$image.Dispose();
	
	return $image_history;
}

Write-Host "Starting watermarking job in $imagesPath on $extensions, $newerThan days or newer...";
$(Get-Date -format T) + " Starting watermarking job in $imagesPath on $extensions, $newerThan days or newer..." >> $logFile;

#Create a list of all files in the directory
$files = Get-ChildItem -Path $imagesPath -Include $extensions | WHERE { $_.CreatedTime -lt ($(Get-Date).AddDays($newerThan * -1)) -and $_.Attributes -ne 'Directory' };
$numFiles = $files.Length;
$i = 1; #Just to know how far it has reached

Write-Host "Found $numFiles files to handle...";
$(Get-Date -format T) + " Found $numFiles files to handle..." >> $logFile;

if ($numFiles -gt 0) {
	foreach ($file in $files) {
		$SourceImage = $file.FullName;
		$TargetImage = $file.FullName; #The same for replacing the old files
		$percentDone = [Math]::Floor(($i / $numFiles) * 100);

		$exifCheck = Read-ImageWaterMark($SourceImage);
		
		if ($exifCheck.Contains($exifIdentifier) -ne $true) {
			Write-Host "$percentDone% Watermarking $SourceImage";
			Write-ImageWaterMark -watermarkImage $watermarkImage -SourceImage $SourceImage -TargetImage $TargetImage;
		}
		else {
			Write-Host "$percentDone% $SourceImage was already watermarked...";
		}
		
		$i++;
	}
}

$(Get-Date -format T) + " Watermarking job is done..." >> $logFile;

exit;
