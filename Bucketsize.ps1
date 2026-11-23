Write-Host(" _____               _    _        ") -f white
Write-Host("|  __ \             | |  (_)       ") -f white
Write-Host("| |  | | _   _  ___ | |_  _  _ __  ") -f white
Write-Host("| |  | || | | |/ __|| __|| || '_ \ ") -f white 
Write-Host("| |__| || |_| |\__ \| |_ | || | | |") -f white
Write-Host("|_____/  \__,_||___/ \__||_||_| |_|") -f white
Write-Host("BucketSize v2.2 by RUNE GOKS0R 2026") -f yellow
Write-Host("(calculates S3 buckets consumption)") -f yellow
Write-Host("")

###### Check if needed modules are in place or install needed ######
if (-not(Get-Module -ListAvailable -Name aws.tools.common )) { Install-Module aws.tools.common } Else { Write-host "aws.tools.common" -f white -nonewline; Write-host " module detected" -f cyan }
if (-not(Get-Module -ListAvailable -Name aws.tools.s3 )) { Install-Module aws.tools.s3 } Else { Write-host "aws.tools.s3 module" -f white -nonewline; Write-host " module detected" -f cyan }
if (-not(Get-Module -ListAvailable -Name ImportExcel )) { Install-Module ImportExcel } Else { Write-host "ImportExcel" -f white -nonewline; Write-host " module detected" -f cyan ; Import-Module ImportExcel }
if (-not(get-awscredential -ProfileName default )) {
	Throw { First you need to activate default AWS credentials with Set-AWSCredentials command. Example: Set-AWSCredential -AccessKey XXXXXX -SecretKey XXXXXX -StoreAs Default }
	} Else {
	Write-host "Default" -f white -nonewline; Write-host " AWS credentials detected" -f cyan
}
Write-Host("`n")

$url = $args[0]
$bu = $args[1]

If ($url -eq $null) {
	If ($bu -eq $null) {
		Write-host "You may pass both url and bucket options like this example:"  -f yellow -nonewline
		Write-host ".\bucketsize.ps1 https://dk-east.s3.dustin.com mybucket"  -f darkyellow
	}
	Write-host "You may pass endpoint url like this example:"  -f yellow -nonewline
	Write-host ".\bucketsize.ps1 https://dk-east.s3.dustin.com"  -f darkyellow
	$url = "https://no-west.s3.dustin.com" 
	Write-host "Scanning S3 default address: "  -f white -nonewline
	Write-host $url -f cyan
} else {
	If ($bu -eq $null) {
		Write-host "You may specify bucket option for singel bucket scan like this example:"  -f yellow -nonewline
		Write-host ".\bucketsize.ps1 https://dk-east.s3.dustin.com mybucket"  -f darkyellow
	}
	Write-host "Scanning S3 address: $url"  -f white -nonewline
	Write-host $url -f cyan
}

###### DataSizeUnitOfMeasure ##### Usage Example: Convert-DataSizeUnitOfMeasure -From TB -To TiB -Value 1024 -Precision 2
function Convert-DataSizeUnitOfMeasure {
  param(            
    [Parameter(position=0,Mandatory=$true,HelpMessage='You must specify the starting unit of measure.')]
    [validateset('b','bit','Nibble','Kbit','Kilobit','Kibibit','Kibit','Mbit','MegaBit','Mebibit','Mibit','Gbit','Gigabit','Gibit','Gibibit','Tbit','Terabit','Tibit','Tebibit','Pbit','Petabit','Pibit','Pebibit','Ebit','Exabit','Eibit','Exbibit','Zbit','Zettabit','Zibit','Zebibit','Ybit','Yottabit','Yibit','Yobibit','B','Byte','KB','Kilobyte','KiB','Kibibyte','MB','Megabyte','MiB','Mebibyte','GB','Gigabyte','GiB','Gibibyte','TB','Terabyte','TiB','Tebibyte','PB','Petabyte','PiB','Pebibyte','EB','Exabyte','EiB','Exbibyte','ZB','Zettabyte','ZiB','Zebibyte','YB','Yottabyte','YiB','Yobibyte',IgnoreCase = $false)]
    [string]$From,
    [Parameter(position=1)]
    [validateset('b','bit','B','Nibble','Byte','KB','Kilobyte','KiB','Kibibyte','MB','Megabyte','MiB','Mebibyte','GB','Gigabyte','GiB','Gibibyte','TB','Terabyte','TiB','Tebibyte','PB','Petabyte','PiB','Pebibyte','EB','Exabyte','EiB','Exbibyte','ZB','Zettabyte','ZiB','Zebibyte','YB','Yottabyte','YiB','Yobibyte',IgnoreCase = $false)]
    [AllowNull()]
    [string]$To,
    [Parameter(position=2,Mandatory=$true,HelpMessage='Please provide a data size to convert')]
    [double]$Value,
    [Parameter(Position=3)]
    [int]$Precision = 4
  )

  # Convert the supplied value to Bytes
  switch -casesensitive ($From) {
    { $_ -in 'b','bit' }         {$value = $value/8 }
    { $_ -in 'Nibble'  }         {$value = $value/2 }
    { $_ -in 'B','Byte' }        {$value = $Value }
    { $_ -in 'Kbit','Kilobit' }  {$value = ($value/8) * [math]::pow(10,3) }
    { $_ -in 'Kibit','Kibibit' } {$value = ($value/8) * [math]::pow(2,10) }
    { $_ -in 'KB','Kilobyte' }   {$value = $Value * [math]::pow(10,3) }
    { $_ -in 'KiB','Kibibyte' }  {$value = $value * [math]::pow(2,10) }
    { $_ -in 'Mbit','Megabit' }  {$value = ($value/8) * [math]::pow(10,6) }
    { $_ -in 'Mibit','Mebibit' } {$value = ($value/8) * [math]::pow(2,20) }
    { $_ -in 'MB','Megabyte' }   {$value = $Value * [math]::pow(10,6) }
    { $_ -in 'MiB','Mebibyte' }  {$value = $value * [math]::pow(2,20) }
    { $_ -in 'Gbit','Gigabit' }  {$value = ($value/8) * [math]::pow(10,9) }
    { $_ -in 'Gibit','Gibibit' } {$value = ($value/8) * [math]::pow(2,30) }
    { $_ -in 'GB','Gigabyte' }   {$value = $Value * [math]::pow(10,9) }
    { $_ -in 'GiB','Gibibyte' }  {$value = $value * [math]::pow(2,30) }
    { $_ -in 'Tbit','Terabit' }  {$value = ($value/8) * [math]::pow(10,12) }
    { $_ -in 'Tibit','Tebibit' } {$value = ($value/8) * [math]::pow(2,40) }
    { $_ -in 'TB','Terabyte' }   {$value = $Value * [math]::pow(10,12) }
    { $_ -in 'TiB','Tebibyte' }  {$value = $value * [math]::pow(2,40) }
    { $_ -in 'Pbit','Petabit' }  {$value = ($value/8) * [math]::pow(10,15) }
    { $_ -in 'Pibit','Pebibit' } {$value = ($value/8) * [math]::pow(2,50) }
    { $_ -in 'PB','Petabyte' }   {$value = $value * [math]::pow(10,15) }
    { $_ -in 'PiB','Pebibyte' }  {$value = $value * [math]::pow(2,50) }
    { $_ -in 'Ebit','Exabit' }   {$value = ($value/8) * [math]::pow(10,18) }
    { $_ -in 'Eibit','Exbibit' } {$value = ($value/8) * [math]::pow(2,60) }
    { $_ -in 'EB','Exabyte' }    {$value = $value * [math]::pow(10,18) }
    { $_ -in 'EiB','Exbibyte' }  {$value = $value * [math]::pow(2,60) }
    { $_ -in 'Zbit','Zettabit' } {$value = ($value/8) * [math]::pow(10,21) }
    { $_ -in 'Zibit','Zebibit' } {$value = ($value/8) * [math]::pow(2,70) }
    { $_ -in 'ZB','Zettabyte' }  {$value = $value * [math]::pow(10,21) }
    { $_ -in 'ZiB','Zebibyte' }  {$value = $value * [math]::pow(2,70) }
    { $_ -in 'Ybit','Yottabit' } {$value = ($value/8) * [math]::pow(10,24) }
    { $_ -in 'Yibit','Yobibit' } {$value = ($value/8) * [math]::pow(2,80) }
    { $_ -in 'YB','Yottabyte' }  {$value = $value * [math]::pow(10,24) }
    { $_ -in 'YiB','Yobibyte' }  {$value = $value * [math]::pow(2,80) }
  }
  
  If ( [string]::IsNullOrWhiteSpace( $To ) )
  { # If no 'To' unit is provided we calcuate the most applicable base 2 unit of measure based upon the number of bytes we've calculated.
    Switch ($Value)
    {
      { $_ -ge 0 -and $_ -lt 1   } { $To = 'b'; break; }
      { $_ -lt [math]::pow(2,10) } { $To = 'B'; break; }
      { $_ -lt [math]::pow(2,20) } { $To = 'KiB'; break; }
      { $_ -lt [math]::pow(2,30) } { $To = 'MiB'; break; }
      { $_ -lt [math]::pow(2,40) } { $To = 'GiB'; break; }
      { $_ -lt [math]::pow(2,50) } { $To = 'TiB'; break; }
      { $_ -lt [math]::pow(2,60) } { $To = 'PiB'; break; }
      { $_ -lt [math]::pow(2,70) } { $To = 'EiB'; break; }
      { $_ -lt [math]::pow(2,80) } { $To = 'ZiB'; break; }
      { $_ -lt [math]::pow(2,90) } { $To = 'YiB'; break; }
      Default { $To = 'YiB'; break; }
    }
  }

  # Convert the number of Bytes to the desired output
  switch -casesensitive ($To) {
    { $_ -in 'b','bit' }         {$value = $value * 8}
    { $_ -in 'Nibble' }          {$value = $value * 2}
    { $_ -in 'B','Byte' }        {$value = $value }
    { $_ -in 'KB','Kilobyte' }   {$value = $Value/[math]::pow(10,3) }
    { $_ -in 'KiB','Kibibyte' }  {$value = $value/[math]::pow(2,10) }
    { $_ -in 'MB','Megabyte' }   {$value = $Value/[math]::pow(10,6) }
    { $_ -in 'MiB','Mebibyte' }  {$value = $Value/[math]::pow(2,20) }
    { $_ -in 'GB','Gigabyte' }   {$value = $Value/[math]::pow(10,9) }
    { $_ -in 'GiB','Gibibyte' }  {$value = $Value/[math]::pow(2,30) }
    { $_ -in 'TB','Terabyte' }   {$Value = $Value/[math]::pow(10,12) }
    { $_ -in 'TiB','Tebibyte' }  {$Value = $Value/[math]::pow(2,40) }
    { $_ -in 'PB','Petabyte' }   {$Value = $Value/[math]::pow(10,15) }
    { $_ -in 'PiB','Pebibyte' }  {$Value = $Value/[math]::pow(2,50) }
    { $_ -in 'EB','Exabyte' }    {$Value = $Value/[math]::pow(10,18) }
    { $_ -in 'EiB','Exbibyte' }  {$Value = $Value/[math]::pow(2,60) }
    { $_ -in 'ZB','Zettabyte' }  {$value = $value/[math]::pow(10,21) }
    { $_ -in 'ZiB','Zebibyte' }  {$value = $value/[math]::pow(2,70) }
    { $_ -in 'YB','Yottabyte' }  {$value = $value/[math]::pow(10,24) }
    { $_ -in 'YiB','Yobibyte' }  {$value = $value/[math]::pow(2,80) }
  }

  $Results = [pscustomobject]@{
    value=[Math]::Round($value,$Precision,[MidPointRounding]::AwayFromZero)
    unit = $To
  }
  return $Results
}

###### Returns a string with no control characters or entities above 0x00FD #######
function Repair-XmlString {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0,ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [string]$String
    )
    #Write-Host "Cleaning string for XML parsing [String: $($String)]"
    $rPattern = "[^\x09\x0A\x0D\x20-\xD7FF\xE000-\xFFFD\x10000\x10FFFF]"
    $cleaned = $String -replace $rPattern, ''
    #Write-Host "Returning parsed string [String cleaned: $($cleaned)]"
    return $cleaned
}

###### Human readability ######
function Get-FriendlySize {
    param($Bytes)
    $sizes='Bytes,KB,MB,GB,TB,PB,EB,ZB' -split ','
    for($i=0; ($Bytes -ge 1kb) -and
        ($i -lt $sizes.Count); $i++) {$Bytes/=1kb}
    $N=1 # change to 2 if 2 decimals and uncomment next line
	if($i -eq 0) {$N=1}
    "{0:N$($N)} {1}" -f $Bytes, $sizes[$i]
}

$totsize = 0
$totfilescount = 0
$totaveragesize = 0
$totscanH = 0
$totscanM = 0
$watch = New-Object System.Diagnostics.Stopwatch

###### List Bucket(s) in account #####
If ($bu -eq $null) {
	Write-Host("")
	Write-Host ("`tBucket not specified. Scanning all buckets in account")  -f white
	Write-Host ("¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨")
	$buckets = Get-s3Bucket -EndpointUrl $url
	$x = 0
	foreach ($bucket in $buckets) {
		$x++
		Write-Host("$x`t") -f yellow -n
		Write-Host($bucket.BucketName) -f cyan
	}
	Write-Host ("`n¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨\ ") -f white -nonewline 
	#Write-Host("--------<| ") -f white -nonewline
	Write-Host("Buckets to scan: ") -f yellow -nonewline
	Write-Host($x) -f darkyellow -n
	Write-Host (" /¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨") -f white
	Write-Host ("                       ¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨") -f white
} else {
	$buckets = Get-s3Bucket -EndpointUrl $url -Bucket $bu
}
Write-Host("`n")

###### Per Bucket Loop ######
foreach ($bucket in $buckets) {
	$char = 0
	$len = 1
	$len = $bucket.BucketName.Length
	$blankspace = 0
	If ($len -lt 25) {
		$len = 24
		$blankspace = $len - $bucket.BucketName.Length
	}
	Write-Host("]-------<| ") -f white -nonewline
	Write-Host($bucket.BucketName) -f cyan
	
	###### Bucket scan #####
	Write-Host("         | ") -f white -nonewline

	$busize=0
	$filescount=0
	$averagesize=0
	$watch.Start()

	###### OLD Bucket Read loop ######
	# $Measurement = Get-S3ObjectV2 $bucket.Bucketname -EndpointUrl $url | Measure-Object -Sum -Average -Property Size

	###### NEW Bucket Read loop ######
		# First batch to establish baseline
		Write-Host("Total (Si) <") -f yellow -nonewline
		$pos = $host.ui.RawUI.get_cursorPosition()
		$batch = (Get-S3Version  -EndpointUrl $url -BucketName $bucket.Bucketname).Versions | Select Key,Size
		If ($batch.Count -ne 0) { 
			$marker = $batch[999].Key
			$batchsize = $batch | Measure-Object -Sum -Property Size
			$busize = $busize + $batchsize.Sum
			# Write-Host ("■ ") -f White -nonewline
			if ($currentcolor -eq "Green") { $currentcolor = "Blue" } else { $currentcolor = "Green" }
			Write-Host ("  ") -nonewline
			Write-Host (Get-FriendlySize($busize)) -f $currentcolor -nonewline
			Write-Host ("   ") -nonewline
			$host.UI.RawUI.set_cursorPosition($Pos)

			# Iterate from 1000 files to the end
			for ($x =2; $batch.Count -gt 999; $x++) {
				try {
						$batch = (Get-S3Version  -EndpointUrl $url -BucketName $bucket.Bucketname -KeyMarker $marker).Versions | Select Size, Key
						$marker = $batch[999].Key
						$batchsize = $batch | Measure-Object -Sum -Property Size
						$busize = $busize + $batchsize.Sum
					}
					catch {
						Write-Host ("Error in Keymarker:",$marker,$batch.Count) -f Red
						Write-Host ("         | Total (Si) <") -f white -nonewline
						$pos = $host.ui.RawUI.get_cursorPosition()
					}
				# Write-Host ("■ ") -f White -nonewline
				if ($currentcolor -eq "Green") { $currentcolor = "Blue" } else { $currentcolor = "Green" }
				Write-Host ("  ") -nonewline
				Write-Host (Get-FriendlySize($busize)) -f $currentcolor -nonewline
				Write-Host ("        ") -nonewline
				$host.UI.RawUI.set_cursorPosition($Pos)
				}	
		###### Calculate Average and Item-count for Bucket #####
		$filescount = ($x * 1000) + $batch.Count
		
		###### Export X GB to CSV and Excel #####
		$SizeGB = [Math]::Round($busize/1000000000,0,[MidPointRounding]::AwayFromZero)

		$averagesize = [Math]::Round($busize/$filescount,0,[MidPointRounding]::AwayFromZero)
		} else {
			$busize = 0
			###### Export 0 to CSV and Excel #####
			$SizeGB = 0 
		}
		# Debug: Write-Host ("SizeGB for EXCEL: ",$SizeGB) -f Red
		$bucket | Add-Member -Membertype NoteProperty -Name SizeGB -value $SizeGB	

	###### Print total for bucket ######
	$pos = $host.ui.RawUI.get_cursorPosition();$pos.X = $pos.X-14;$host.UI.RawUI.set_cursorPosition($Pos)
	Write-Host("| Total (Si)  : ") -f yellow -nonewline
	Write-Host(Get-FriendlySize($busize)) -f DarkYellow -nonewline 
	Write-Host("      ")
	
	#Write-Host("         | ") -f white -nonewline
	#Write-Host("Total TiB   : ") -f yellow -nonewline	
	#$si = Convert-DataSizeUnitOfMeasure -From Byte -Value $busize -To TiB -Precision 1
	#Write-Host($si.value,$si.unit) -f DarkYellow
	
	Write-Host("         | ") -f white -nonewline
	Write-Host("Total TB    : ") -f yellow -nonewline	
	$si = Convert-DataSizeUnitOfMeasure -From Byte -Value $busize -To TB -Precision 1
	Write-Host($si.value,$si.unit) -f Green
	
	#If ($busize -lt 1024) {Convert-DataSizeUnitOfMeasure -From Byte -Value $totsize -To KiB -Precision 0;Write-Host " KiB"}
	#If ($busize -ge 1024) {Convert-DataSizeUnitOfMeasure -From Byte -Value $totsize -To MiB -Precision 0;Write-Host " MiB"}
	
	# Write-Host(Get-FriendlySize($busize)) -f DarkYellow
	Write-Host("         | ") -f white -nonewline
	Write-Host("Total Files : ") -f yellow -nonewline
	Write-Host($filescount) -f DarkYellow
	Write-Host("         | ") -f white -nonewline
	Write-Host("Average Size: ") -f yellow -nonewline
	Write-Host(Get-FriendlySize($averagesize)) -f DarkYellow
	# Write-Host($AverageMB,"MB") -f DarkYellow
	Write-Host("         | ") -f white -nonewline
	Write-Host("Elapsed Time: ") -f yellow -nonewline
	Write-Host($watch.Elapsed.Hours) -f DarkYellow -nonewline
	$totscanH = $totscanH + $watch.Elapsed.Hours
	Write-Host("h") -f DarkYellow -nonewline
	Write-Host($watch.Elapsed.Minutes) -f DarkYellow -nonewline
	$totscanM = $totscanM + $watch.Elapsed.Minutes
	Write-Host("m") -f DarkYellow
	Write-Host("         '") -f white -nonewline
	for ($char =0; $char -le $len; $char++) { Write-Host("=") -f white -nonewline }
	Write-Host("='") -f white -nonewline
	$pos = $host.ui.RawUI.get_cursorPosition();$pos.Y = $pos.Y-6;$pos.X = $pos.X-1;$host.UI.RawUI.set_cursorPosition($Pos);Write-Host("|>") -f white -nonewline
	$pos = $host.ui.RawUI.get_cursorPosition();$pos.X = $pos.X-2;$pos.Y = $pos.Y+1;$host.UI.RawUI.set_cursorPosition($Pos);Write-Host("|") -f white -nonewline
	$pos = $host.ui.RawUI.get_cursorPosition();$pos.X = $pos.X-1;$pos.Y = $pos.Y+1;$host.UI.RawUI.set_cursorPosition($Pos);Write-Host("|") -f white -nonewline
	$pos = $host.ui.RawUI.get_cursorPosition();$pos.X = $pos.X-1;$pos.Y = $pos.Y+1;$host.UI.RawUI.set_cursorPosition($Pos);Write-Host("|") -f white -nonewline
	$pos = $host.ui.RawUI.get_cursorPosition();$pos.X = $pos.X-1;$pos.Y = $pos.Y+1;$host.UI.RawUI.set_cursorPosition($Pos);Write-Host("|") -f white -nonewline
	$pos = $host.ui.RawUI.get_cursorPosition();$pos.X = $pos.X-1;$pos.Y = $pos.Y+1;$host.UI.RawUI.set_cursorPosition($Pos);Write-Host("|") -f white -nonewline
	Write-Host("`n","`n")
	
	$watch.Reset()
	$totsize = $totsize + $busize
	if ($totaveragesize -eq 0) {$totaveragesize = ($totaveragesize + $averagesize)} else {$totaveragesize = ($totaveragesize + $averagesize) / 2}
	$totfilescount = $totfilescount + $filescount
}
$totaveragesize = [Math]::Round($totaveragesize,0,[MidPointRounding]::AwayFromZero)


###### Print total for account ######
If ($bu -eq $null) {
	$tots = Convert-DataSizeUnitOfMeasure -From Byte -Value $totsize -To TiB -Precision 1
	#Write-Host("Total Account Size in bytes: ") -f yellow -nonewline;Write-Host($totsize) -f Green
	$totgib = Convert-DataSizeUnitOfMeasure -From Byte -Value $totsize -To TB -Precision 1
	$totgibtxt = $totgib.value,$totgib.unit
	$tota = Get-FriendlySize($totaveragesize)
	$len = 0
	$extrahours = 0
	
	###### Converting extra minutes to hours ######
	if ($totscanM -lt 60) {
		$extrahours = ($totscanM - 60) / 60
		$extrahours = [Math]::Round($extrahours,0,[MidPointRounding]::AwayFromZero)
		$totscanH = $totscanH + $extrahours
		}
	If ($totscanH -lt 0) {$totscanH = 0}
		
	Write-Host("`n")
	#Write-Host("         .") -f white -nonewline
	#$len = "$tots.value".Length + "$tots.unit".Length - 27;for ($char =1; $char -le $len; $char++) { Write-Host("=") -f white -nonewline }
	#Write-Host("=.") -f white
	
	Write-Host("]-------<| ") -f white -nonewline

	Write-Host("Total Account (Si): ") -f yellow -nonewline
	Write-Host($tots.value,$tots.unit) -f DarkYellow -nonewline
	#for ($char =1; $char -le $len; $char++) { Write-Host(" ") -f white -nonewline }

	#Write-Host("   |") -f white
	Write-Host("`n]-------<| ") -f white -nonewline

	Write-Host("Total Account TB  : ") -f yellow -nonewline
	Write-Host($totgib.value,$totgib.unit) -f Green -nonewline
	$lengib1 = ($totgib.value | Measure-Object -Character).Characters
	$lengib2 = ($totgib.unit | Measure-Object -Character).Characters
	$lengib = 10 - $lengib1 - $lengib2 ;for ($char =1; $char -le $lengib; $char++) { Write-Host(" ") -f white -nonewline }

	#Write-Host(" |") -f white
	Write-Host("`n]-------<| ") -f white -nonewline

	Write-Host("Total Files Count : ") -f yellow -nonewline
	Write-Host($totfilescount) -f DarkYellow -nonewline
	$lencount = 11 - "$totfilescount".Length;for ($char =1; $char -le $lencount; $char++) { Write-Host(" ") -f white -nonewline }

	#Write-Host(" |") -f white
	Write-Host("`n]-------<| ") -f white -nonewline

	Write-Host("Total Scan Time   : ") -f yellow -nonewline
	Write-Host($totscanH) -f DarkYellow -nonewline
	Write-Host("h") -f DarkYellow -nonewline
	Write-Host($totscanM) -f DarkYellow -nonewline
	Write-Host("m") -f DarkYellow -nonewline
	#$lenscan = 9 - "$totscanH".Length - "$totscanM".Length;for ($char =1; $char -le $lenscan; $char++) { Write-Host(" ") -f white -nonewline }
	#Write-Host(" |") -f white
	
	Write-Host("`n]-------<| ") -f white -nonewline
	Write-Host("Total Average Size: ") -f yellow -nonewline
	Write-Host($tota) -f DarkYellow -nonewline
	#$lenavg = 11 - "$tota".Length;for ($char =1; $char -le $lenavg; $char++) { Write-Host(" ") -f white -nonewline }
	#Write-Host(" |") -f white
	
	#Write-Host("         '") -f white -nonewline
	#for ($char =0; $char -le $len; $char++) { Write-Host("=") -f white -nonewline }
	#Write-Host("'") -f white
	
	Write-Host("`n")
	# if (Test-Path '.\buckets.xlsx') { Remove-Item '.\buckets.xlsx' }
	###### Export to CSV ######
	Write-Host("Exporting size-table to: ") -f yellow -nonewline
	$TimeMarker = Get-Date -Format "ddMMyyyy_HHmm"
	$buckets | Export-CSV -Path .\buckets-$TimeMarker.csv
	Write-Host(".\buckets-$TimeMarker.csv") -f DarkYellow
	###### Export to Excel ######
	$buckets | Where BucketName | Export-Excel -Path .\buckets-$TimeMarker.xlsx -WorksheetName Buckets -ChartType PieExploded3D -IncludePivotChart -IncludePivotTable -PivotRows BucketName -PivotData @{SizeGB='sum'}
	Write-Host("                         .\buckets-$TimeMarker.xlsx") -f DarkYellow
	Write-Host("")
}