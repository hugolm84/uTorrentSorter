Param(
	[parameter(Mandatory = $true)][String]$TORRENT_NAME, 
	[parameter(Mandatory = $true)][String]$TORRENT_DIR, 
	## EDIT ->>
	# Set the baseDir for your sorting and unpacking	
	[String]$BASE_MOVIE_DIR = "F:\media\movies\", 
	[String]$BASE_TV_DIR = "F:\media\tv\", 
	# Want to enable loggin if there is errors?	
	# If Set to NO, it will only be outputted in console
	[String]$LOGG = "YES",
	# Set the correct path to winRar	
	[String]$WINRAR = "C:\Program Files\WinRAR\"
	## <<- END OF EDIT
	)

# Set scriptPath first
$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition;

function Write-Log{
	Param([string] $WriteLog)
	$date = Get-Date;
	$loggStr = "$date : $WriteLog"
	
	if($LOGG -eq "YES"){
		$loggStr >> $scriptPath\logFile.log
	}else{
		write-output $WriteLog;
	}
}

# Test crucial dirs and drives
$driveOnMovieBase = (New-Object System.IO.DriveInfo($BASE_MOVIE_DIR)).DriveType -ne 'NoRootDirectory'
$driveOnTvBase = (New-Object System.IO.DriveInfo($BASE_TV_DIR)).DriveType -ne 'NoRootDirectory'
	if(-not($driveOnMovieBase)){ Write-Log "No Movie Volume!" break  }
	if(-not($driveOnTvBase)){ Write-Log "No Tv Volume!" break }     
	if(-not(Test-Path $WINRAR)){ Write-Log "Can not find WinRar in path!" break }
# Test End

# Add paths to env:path
$env:Path = $env:Path + ";$scriptPath\bin";		
$env:Path = $env:Path + ";$WINRAR";	

# Set dirs, create if not exist
# Note: You can edit the names of these, but that would break the script. 
# To edit or add, you need to edit/add regex in getData
$unpackToMovieDir = "${BASE_MOVIE_DIR}unpacked\$TORRENT_NAME";
$unpackToTvDir = "${BASE_TV_DIR}unpacked\$TORRENT_NAME";
$movieDirs = ("unpacked", "sorted\genre", "sorted\rating", "sorted\year", "sorted\title", "sorted\metascore");
$tvDirs = ("unpacked", "sorted\genre", "sorted\title");

# Create movie and tv dirs
$dirs = @(foreach($dir in $movieDirs){ 
			Join-Path $BASE_MOVIE_DIR $dir 
		}) + 
		@(foreach($dir in $tvDirs){ 
			Join-Path $BASE_TV_DIR $dir 
		})
	md $dirs -Force -ErrorAction Stop | Out-Null

	if( (Test-Path $dirs -type container) -contains $false ) { 
		throw "Some folders weren't created because there were files in the way, I can't figure out what to do about that" 
	}


# check the NFO for imdb och tvrage link
function Test-NFO{
	if(Test-Path $TORRENT_DIR){
		$nfoItem = get-childitem $TORRENT_DIR -recurse | where {$_.extension -eq ".nfo"}
		if($nfoItem){
			$isMovie = $nfoItem | Select-String "http://www.imdb.com/([\S]*)" | select -exp Matches | select -exp value
			$isTV = $nfoItem | Select-String "http://www.tvrage.com/([\S]*)" | select -exp Matches | select -exp value
			if($isMovie){ 
				$url = $isMovie;
				$type = "Movie";
			}
			elseif($isTv){ 
				$url = "$isTv";
				$type = "Tv";
			}
			return $ret = @{url = $url; type = $type;}
		}
	}
}

# Set the path for TV or Movie
function Set-Path{
	Param($ret)
	if($ret['type'] -eq "Tv"){
		return $unpackToTvDir;
	}
	if($ret['type'] -eq "Movie"){
		return $unpackToMovieDir;
	}
	if(!$ret['type']){
		return "Error";
	}
}


# Unrars every *.rar file recursivly. that is, CD1 CD2, subs and so on
function Extract-Rar{
	Param($unpackDestination)
	if(Test-Path $TORRENT_DIR){
		if(-not($unpackDestination -eq "Error")){
			
			if(-not(Test-Path $unpackDestination)){
				md $unpackDestination;
			}
			$rarItems = Get-ChildItem $TORRENT_DIR -recurse -Filter *.rar | Select-Object -ExpandProperty FullName
			foreach($item in $rarItems){
				& unrar e -o- $item $unpackDestination;
			}
			# Unrars sub item and then removes the package. This is in the unpacked destination folder and not Original folder
			$subItems = Get-ChildItem $unpackDestination -recurse -Filter *.rar | Select-Object -ExpandProperty FullName
			if($subItems){
				foreach($item in $subItems){
					& unrar e -o- $item $unpackDestination;
					ri $item;
				}
			}
		}
	}
}

# Get data from IMDB, based on nfo URL
function Get-IMDb-Data{ 
    param([string] $url) 
    $wc = New-Object System.Net.WebClient 
    $data = $wc.downloadstring($url) 
    # $title = [regex] '(?<=<title>)([\S\s]*?)(?=</title>)' 
	$title = ([regex]'(?<=<h1 class="header">)([\S\s]*?)(?=<span>)').Match($data).value.trim();
	$year = ([regex]'(?<=<span>[(]<a href="/year/)(([\S\s]*?)+)(?=/">)').Match($data).value.trim().TrimEnd();
	$rating = ([regex] '(?<=<span class="rating-rating">)(([\S\s]*?)+)(?=<span>)').Match($data).value.trim();
	$genre1, $genre2, $genre3 = ([regex]'(?<=<a href="/genre/)(([\S\s]*?)+)(?=")').matches($data) | foreach {$_.Groups[1].Value}
	$metascore = ([regex]'(?<=<span class="nobr">Metascore\S\s*<strong>)([\S\s]*?)(?=</strong>)').Match($data).value.trim();
	$byTitle = $title.substring(0,1);
	return $ret = @{title = $title;
					bytitle = "${BASE_MOVIE_DIR}sorted\title\$byTitle\$title";
					year = "${BASE_MOVIE_DIR}sorted\year\$year";
					rating = "${BASE_MOVIE_DIR}sorted\rating\$rating"; 
					metascore = "${BASE_MOVIE_DIR}sorted\metascore\$metascore";
					genre1 = "${BASE_MOVIE_DIR}sorted\genre\$genre1";
					genre2 = "${BASE_MOVIE_DIR}sorted\genre\$genre2";
					genre3 = "${BASE_MOVIE_DIR}sorted\genre\$genre3"
					}; 

	}

# Get data from tvRage, based on nfo URL
function Get-TVrage-Data { 
	Param($url)
	$season = $TORRENT_NAME -replace '.*s(.*)e.*','$1';
	$title, $genres = (get-webfile http://services.tvrage.com/tools/quickinfo.php?show=$url -passthru ) -split "`n" | select -index 1,13
	$title = $title -replace "Show Name@" -replace ":" -replace " ", "_";
	$genres = $genres -replace "Genres@", "";
	$genre1,$genre2,$genre3 = $genres.split("|")|%{$_.trim()}
	return $ret = @{title = $title;
					season = "${BASE_TV_DIR}sorted\title\$title\season_$season";
					genre1 = "${BASE_TV_DIR}sorted\genre\$genre1";
					genre2 = "${BASE_TV_DIR}sorted\genre\$genre2";
					genre3 = "${BASE_TV_DIR}sorted\genre\$genre3";
					} 
	}

# Create symlinks for the torrent, this will link to the unpacked dir	
function Create-SymLinks{
	Param($path, $data)
	if($data){
		foreach ($key in $data.keys -ne 'title') {
			if(-not (Test-Path $data.$key)){
				md $data.$key; 
			}

			$dir = Join-Path $data.$key $TORRENT_NAME
			mklnk -a $path c:\windows\explorer.exe $dir;
		}
	}
}

## Get-WebFile (aka wget for PowerShell)
function Get-WebFile {
[CmdletBinding()]
   param(
      [Parameter(Mandatory=$true,Position=0)]
      [string]$Url # = (Read-Host "The URL to download")
   ,
      [string]$FileName
   ,
      [switch]$Passthru,
      [switch]$Quiet,
      [string]$UserAgent = "PoshCode/$($PoshCode.ScriptVersion)"      
   )

   Write-Verbose "Downloading '$url'"

   $request = [System.Net.HttpWebRequest]::Create($url);
   $request.UserAgent = $(
         "{0} (PowerShell {1}; .NET CLR {2}; {3}; http://PoshCode.org)" -f $UserAgent, 
         $(if($Host.Version){$Host.Version}else{"1.0"}),
         [Environment]::Version,
         [Environment]::OSVersion.ToString().Replace("Microsoft Windows ", "Win")
      ) 
   if($request.Proxy -ne $null) {
      $request.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
   }

   try {
      $res = $request.GetResponse();
   } catch [System.Net.WebException] { 
      Write-Error $_.Exception -Category ResourceUnavailable
      return
   }
 
   if((Test-Path variable:res) -and $res.StatusCode -eq 200) {
      if($fileName -and !(Split-Path $fileName)) {
         $fileName = Join-Path (Convert-Path (Get-Location -PSProvider "FileSystem")) $fileName
      }
      elseif((!$Passthru -and !$fileName) -or ($fileName -and (Test-Path -PathType "Container" $fileName)))
      {
         [string]$fileName = ([regex]'(?i)filename=(.*)$').Match( $res.Headers["Content-Disposition"] ).Groups[1].Value
         $fileName = $fileName.trim("\/""'")
         
         $ofs = ""
         $fileName = [Regex]::Replace($fileName, "[$([Regex]::Escape(""$([System.IO.Path]::GetInvalidPathChars())$([IO.Path]::AltDirectorySeparatorChar)$([IO.Path]::DirectorySeparatorChar)""))]", "_")
         $ofs = " "
         
         if(!$fileName) {
            $fileName = $res.ResponseUri.Segments[-1]
            $fileName = $fileName.trim("\/")
            if(!$fileName) { 
               $fileName = Read-Host "Please provide a file name"
            }
            $fileName = $fileName.trim("\/")
            if(!([IO.FileInfo]$fileName).Extension) {
               $fileName = $fileName + "." + $res.ContentType.Split(";")[0].Split("/")[1]
            }
         }
         $fileName = Join-Path (Convert-Path (Get-Location -PSProvider "FileSystem")) $fileName
      }
      if($Passthru) {
         $encoding = [System.Text.Encoding]::GetEncoding( $res.CharacterSet )
         [string]$output = ""
      }
 
      [int]$goal = $res.ContentLength
      $reader = $res.GetResponseStream()
      if($fileName) {
         $writer = new-object System.IO.FileStream $fileName, "Create"
      }
      [byte[]]$buffer = new-object byte[] 4096
      [int]$total = [int]$count = 0
      do
      {
         $count = $reader.Read($buffer, 0, $buffer.Length);
         if($fileName) {
            $writer.Write($buffer, 0, $count);
         } 
         if($Passthru){
            $output += $encoding.GetString($buffer,0,$count)
         } elseif(!$quiet) {
            $total += $count
            if($goal -gt 0) {
               Write-Progress "Downloading $url" "Saving $total of $goal" -id 0 -percentComplete (($total/$goal)*100)
            } else {
               Write-Progress "Downloading $url" "Saving $total bytes..." -id 0
            }
         }
      } while ($count -gt 0)
      
      $reader.Close()
      if($fileName) {
         $writer.Flush()
         $writer.Close()
      }
      if($Passthru){
         $output
      }
   }
   if(Test-Path variable:res) { $res.Close(); }
}

################################################################
################################################################

# Now we can Execute this, Instantiate, if you will
$nfo = Test-NFO;
if(!$nfo){
	Write-Log "Could not find an NFO!";
}elseif($nfo){
	if(!$nfo['type']){
		Write-Log "Could not determine type!";
	}elseif(!$nfo['url']){
		Write-Log "Could not find an URL!";
	}else{
	
		# First, check the type from nfo and set path
		if($nfo['type'] -eq "Movie"){
			$path = Set-Path $nfo;
			$typeInfo = Get-IMDb-Data $nfo['url'];
		}elseif($nfo['type'] -eq "Tv"){
			$path = Set-Path $nfo;
			$typeInfo = Get-TVrage-Data $nfo['url'];
		}
		# Will now unpack to $path
		Extract-Rar $path;
		# And create symlinks
		Create-SymLinks $path $typeInfo;
	}
}
<#
.SYNOPSIS 
	Automates the logic of sorting downloaded movie and tv torrents.

.DESCRIPTION
 .
	@TYPE
		
		POWERSHELL 2.0									
	
	@WHAT:
		
		uTORRENT onComplete run program 	
	
	@AUTHOR: 
		
		uffepuffe				
	
	@DATE:
		
		2011	
	
	@USAGE:
		
		ADD THIS in uTorrent			
		
			powershell.exe C:\PathToThisSCRIPT.ps1 -TORRENT_NAME '%N' -TORRENT_DIR '%D'	
		
		OR THIS if you want to read the output
			
			powershell.exe C:\PathToThisSCRIPT.ps1 -noexit -TORRENT_NAME '%N' -TORRENT_DIR '%D'	
		
		IF you want to specify the Paths directly you can do
			
			powershell.exe C:\PathToThisSCRIPT.ps1 -TORRENT_NAME '%N' -TORRENT_DIR '%D'-BASE_MOVIE_DIR "C:\PATH"
		
		IF you want to disable/enable loggin
			
			powershell.exe C:\PathToThisSCRIPT.ps1 -TORRENT_NAME '%N' -TORRENT_DIR '%D'-LOGG "NO/YES"
		
		IF you want to alter the winRAR path
			
			powershell.exe C:\PathToThisSCRIPT.ps1 -TORRENT_NAME '%N' -TORRENT_DIR '%D'-WINRAR "D:\PATH_TO_WINRAR"			
	
	@TODO: 
		
		Tv Episodes that are DVDrips probebly has an IMDB link and not a
		TvRage link, wich means that the TV Show DVDrip is treated as a 
		Movie. This should be fixed later somehow.
	
	@DEPENDENCIES: 
		
		\bin\mklnk.exe (included), winRAR (not included), uTorrent	
		mklnk is Copyright (c) 2005-2006 Ross Smith II (http://smithii.com) All Rights Reserved
	
	
	@INFORMATION:
		
		So, almost everybody run Windows 7 or Vista these days. If you do, Powershell is inlcuded and to test that
		you have it installed, open a command promt (run -> cmd) and type powershell $psversiontable
		
		Output:
		
		C:\Windows\System32>powershell $psversiontable
		
		Name                           Value
		----                           -----
		PSVersion                      2.0
		
		So, if this works, and PSVersion = 2.0 we are good to go.
		
		THEN:
		
		Type> Get-ExecutionPolicy
		Output> Restricted
		
		If the output says Restricted, it means that the default value is still set. Change this by doing
		
		Type> Set-ExecutionPolicy RemoteSigned
		Read more here: http://technet.microsoft.com/en-us/library/ee176949.aspx
		
		Else:
		
		goto http://support.microsoft.com/kb/968929
		scroll to
		Windows Management Framework Core (WinRM 2.0 and Windows PowerShell 2.0)
		Download for your windows version
		Install
		OR: 
		Do a system update. It should be included.

.LINK
	.
	@SCREENCASTs:
		@How it operates
			http://www.swfcabin.com/open/1296763224
		@How to set Execution Policy
			http://www.swfcabin.com/open/1296731645	
			

#>