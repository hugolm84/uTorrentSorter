<#
.SYNOPSIS 
	Automates the logic of sorting downloaded movie and tv torrents.
.DESCRIPTION
	Type:
		POWERSHELL 2.0									
	What:
		uTORRENT onComplete run program 	
	Author:
		ulf@flashback				
	Usage:
		ADD THIS in uTorrent	
		
			> powershell.exe C:\PathToThisSCRIPT.ps1 -TORRENT_NAME '%N' -TORRENT_DIR '%D'	
			---------------------------------------------------------------------------
		OR THIS if you want to read the output
		
			> powershell.exe C:\PathToThisSCRIPT.ps1 -noexit -TORRENT_NAME '%N' -TORRENT_DIR '%D'	
			-----------------------------------------------------------------------------------
		IF you want to specify the Paths directly you can do
		
			> powershell.exe C:\PathToThisSCRIPT.ps1 -TORRENT_NAME '%N' -TORRENT_DIR '%D'-BASE_MOVIE_DIR "C:\PATH"
			------------------------------------------------------------------------------------------------------
		IF you want to disable/enable loggin
		
			> powershell.exe C:\PathToThisSCRIPT.ps1 -TORRENT_NAME '%N' -TORRENT_DIR '%D'-LOGG "NO/YES"
			-------------------------------------------------------------------------------------------
		IF you want to alter the winRAR path
		
			> powershell.exe C:\PathToThisSCRIPT.ps1 -TORRENT_NAME '%N' -TORRENT_DIR '%D'-WINRAR "D:\PATH_TO_WINRAR"
			--------------------------------------------------------------------------------------------------------		
		To get more information about the PARAMS
		
			> powershell get-help .\PathTOScript.ps1 -detailed
			--------------------------------------------------
	Todo: 
	
		Tv Episodes that are DVDrips probebly has an IMDB link and not a
		TvRage link, wich means that the TV Show DVDrip is treated as a 
		Movie. This should be fixed later somehow.
	
	Dependencies: 
		
		\bin\mklnk.exe (included), winRAR (not included), uTorrent	
		mklnk is Copyright (c) 2005-2006 Ross Smith II (http://smithii.com) All Rights Reserved
	
	
	Information:
		
		So, almost everybody run Windows 7 or Vista these days. If you do, Powershell is inlcuded and to test that
		you have it installed, open a command promt (run -> cmd) and type powershell $psversiontable
		
		Output:
		
		C:\Windows\System32>powershell $psversiontable
		
		Name                           Value
		----                           -----
		PSVersion                      2.0
		
		
		So, if this works, and PSVersion = 2.0 we are good to go.
		
		Then:
		
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
	Github:
		https://github.com/mline/uTorrentSorter
	Screencasts:
		How it operates
			http://www.swfcabin.com/open/1296763224
			http://www.swfcabin.com/open/1296814280
		How to set Execution Policy
			http://www.swfcabin.com/open/1296731645	
	
#>
Param(
	# This is the Name Parameter sent from the Torrent Client
	[parameter(Mandatory = $true)][String]$TORRENT_NAME, 
	# This is the Dir Parameter sent from the Torrent Client
	[parameter(Mandatory = $true)][String]$TORRENT_DIR, 
	# Set the baseDir for your Movies	
	[String]$BASE_MOVIE_DIR = "F:\media\movies\", 
	# Set the baseDir for your TV shows
	[String]$BASE_TV_DIR = "F:\media\tv\", 
	# Set the baseDir for your Music
	[String]$BASE_MUSIC_DIR = "F:\media\music\",
	# Option to unpack or not
	[String]$UNPACK = "YES",
	# Set the Unpack dir for your Movies
	[String]$UNPACK_MOVIE_TO_DIR = "${BASE_MOVIE_DIR}unpacked\$TORRENT_NAME",
	# Set the Unpack dir for your Tv shows
	[String]$UNPACK_TV_TO_DIR = "${BASE_TV_DIR}unpacked\$TORRENT_NAME",
	# If Set to NO, it will only be outputted in console
	[String]$LOGG = "NO",
	# Set the correct path to winRar	
	[String]$WINRAR = "C:\Program Files\WinRAR\"
	)

# Set scriptPath first
$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition;

# Write-Log function
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
$driveOnMusicBase = (New-Object System.IO.DriveInfo($BASE_MUSIC_DIR)).DriveType -ne 'NoRootDirectory'
	if(-not($driveOnMovieBase)){ Write-Log "No Movie Volume!" break  }
	if(-not($driveOnTvBase)){ Write-Log "No Tv Volume!" break }     
	if(-not($driveOnTvBase)){ Write-Log "No Music Volume!" break }     
	if(-not(Test-Path $WINRAR)){ Write-Log "Can not find WinRar in path!" break }
# Test End

# Add paths to env:path
$env:Path = $env:Path + ";$scriptPath\bin"	
$env:Path = $env:Path + ";$WINRAR"


# Set dirs, create if not exist
# Note: You can edit the names of these, but that would break the script. 
# To edit or add, you need to edit/add regex in getData

$movieDirs = ("unpacked", "by_genre", "by_rating", "by_year", "by_title", "by_metascore", "by_group")
$tvDirs = ("unpacked", "by_genre", "by_title", "by_group")

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
	
# Test for an mp3 file
function Test-MP3{
	if(Test-Path $TORRENT_DIR){
		$mp3Items = Get-ChildItem $TORRENT_DIR -Recurse -Filter *.mp3 | Group-Object DirectoryName | ForEach-Object { $_.Group | Select-Object DirectoryName, FullName -First 1}
		if($mp3Items){
			return $mp3Items
		}
	}
}

# Add the symlink to the mp3 archive
function Add-MP3ToArchive{
	Param($mp3Items)
	# Load the id3 lib
	[Reflection.Assembly]::LoadFrom( "$scriptPath\bin\taglib-sharp.dll") 
	# Maybe we have multiple mp3 results, hence the foreach
	foreach($mp3 in $mp3Items){
		# Read the ID3 tags
		$media = [TagLib.File]::Create($mp3.FullName)
		# Foreach tag
		foreach($tag in $media.Tag){ 
			# Check if the curcial tags exists
			if(-not($tag.Title) -OR (-not($tag.Album)) -OR (-not($tag.Year)) -OR (-not($tag.FirstArtist)) -OR (-not($tag.FirstGenre))){ 
				throw "I need a Title, Album, Year, Artist, Genre to work!"
			}
			# Append the vars 
			else{ 
				$byArtist = $tag.FirstArtist.substring(0,1)
				$artist = $tag.FirstArtist -replace ":" -replace " ", "_" -replace "'"
				$album = $tag.Album -replace ":" -replace " ", "_" -replace "'"
				$genre = $tag.FirstGenre -replace ":" -replace " ", "_" -replace "'"
				$year = $tag.Year -replace ":" -replace " ", "_" -replace "'"
				# Get the what group released this
				$group = ([regex]'([a-zA-Z]*)$').Match($mp3.DirectoryName).value.trim()
				# rlsName, this is because maybe we have multiple rlsNames in one download
				$rlsName = ([regex]'[^\\]*$').Match($mp3.DirectoryName).value.trim()
				# Archive A-Z
				# It is here we will place the symlink $album
				$archiveToDirs = @("${BASE_MUSIC_DIR}sorted\by_genre\$genre",
									"${BASE_MUSIC_DIR}sorted\by_year\$year",
									"${BASE_MUSIC_DIR}sorted\by_artist\$byArtist\$artist",
									"${BASE_MUSIC_DIR}sorted\by_group\$group"
									)
				# Forach dir create them						
				foreach($dir in $archiveToDirs){					
					md $dir -Force -ErrorAction Stop | Out-Null
					mklnk -a $mp3.DirectoryName c:\windows\explorer.exe "$dir\$rlsName"
					write-output "Created dirs for:" 
					write-output "$dir\$rlsName"
		
				}
			}
		}
	}
}

# check the NFO for imdb och tvrage link
function Test-NFO{
	# Does the torrent dir exist?
	if(Test-Path $TORRENT_DIR){
		# Find an nfo
		$nfoItem = get-childitem $TORRENT_DIR -recurse | where {$_.extension -eq ".nfo"}
		if($nfoItem){
			# Now check wheter its a movie or tv
			$isMovie = $nfoItem | Select-String "http://www.imdb.com/([\S]*)" | select -exp Matches | select -exp value
			$isTV = $nfoItem | Select-String "http://www.tvrage.com/([\S]*)" | select -exp Matches | select -exp value
			# Set the type
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

# Unrars every *.rar file recursivly. that is, CD1 CD2, subs and so on
function Extract-Rar{
	Param($unpackDestination)
	# Does the path exist?
	if(Test-Path $TORRENT_DIR){
		if(-not($unpackDestination -eq "Error")){
			# Create destination
			if(-not(Test-Path $unpackDestination)){
				md $unpackDestination;
			}
			$rarItems = Get-ChildItem $TORRENT_DIR -recurse -Filter *.rar | Select-Object -ExpandProperty FullName
			# Unrar the files
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

# Set the path for TV or Movie
function Set-Path{
	Param($ret)
	if($UNPACK -eq "YES"){
		if($ret['type'] -eq "Tv"){
			return $UNPACK_TV_TO_DIR;
		}
		if($ret['type'] -eq "Movie"){
			return $UNPACK_MOVIE_TO_DIR;
		}
		if(!$ret['type']){
			return "Error";
		}
	}else{ 
		return $TORRENT_DIR
	}	
}

# Get data from IMDB, based on nfo URL
function Get-IMDb-Data{ 
    param([string] $url) 
	# Create a new webclient
	$wc = New-Object System.Net.WebClient 
    $data = $wc.downloadstring($url) 
	# Extract data
	$title = ([regex]'(?<=<h1 class="header">)([\S\s]*?)(?=<span>)').Match($data).value.trim();
	$year = ([regex]'(?<=<span>[(]<a href="/year/)(([\S\s]*?)+)(?=/">)').Match($data).value.trim().TrimEnd();
	$rating = ([regex] '(?<=<span class="rating-rating">)(([\S\s]*?)+)(?=<span>)').Match($data).value.trim();
	$genre1, $genre2, $genre3 = ([regex]'(?<=<a href="/genre/)(([\S\s]*?)+)(?=")').matches($data) | foreach {$_.Groups[1].Value}
	$metascore = ([regex]'(?<=<span class="nobr">Metascore\S\s*<strong>)([\S\s]*?)(?=</strong>)').Match($data).value.trim();
	$byTitle = $title.substring(0,1);
	$group = ([regex]'([a-zA-Z]*)$').Match($TORRENT_NAME).value.trim();
	return $ret = @{title = $title;
					bytitle = "${BASE_MOVIE_DIR}by_title\$byTitle\$title";
					year = "${BASE_MOVIE_DIR}by_year\$year";
					rating = "${BASE_MOVIE_DIR}by_rating\$rating"; 
					metascore = "${BASE_MOVIE_DIR}by_metascore\$metascore";
					genre1 = "${BASE_MOVIE_DIR}by_genre\$genre1";
					genre2 = "${BASE_MOVIE_DIR}by_genre\$genre2";
					genre3 = "${BASE_MOVIE_DIR}by_genre\$genre3";
					groupe = "${BASE_MOVIE_DIR}by_group\$group";
					}; 

	}

# Get data from tvRage, based on nfo URL
function Get-TVrage-Data { 
	Param([String]$url)
	$season = $TORRENT_NAME -replace '.*s(.*)e.*','$1';
	$title, $genres = (get-webfile http://services.tvrage.com/tools/quickinfo.php?show=$url -passthru ) -split "`n" | select -index 1,13
	$title = $title -replace "Show Name@" -replace ":" -replace " ", "_";
	$genres = $genres -replace "Genres@", "";
	$genre1,$genre2,$genre3 = $genres.split("|")|%{$_.trim()}
	$group = ([regex]'([a-zA-Z]*)$').Match($TORRENT_NAME).value.trim();
	return $ret = @{title = $title;
					season = "${BASE_TV_DIR}by_title\$title\season_$season";
					genre1 = "${BASE_TV_DIR}by_genre\$genre1";
					genre2 = "${BASE_TV_DIR}by_genre\$genre2";
					genre3 = "${BASE_TV_DIR}by_genre\$genre3";
					groupe = "${BASE_TV_DIR}by_group\$group"
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
         "{0} (PowerShell {1}; .NET CLR {2}; {3}; http://imTryingToSort.org)" -f $UserAgent, 
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
$mp3 = Test-MP3
if($mp3){
	Add-MP3ToArchive $mp3
}else{
	$nfo = Test-NFO
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
				$path = Set-Path $nfo
				$typeInfo = Get-IMDb-Data $nfo['url']
			}elseif($nfo['type'] -eq "Tv"){
				$path = Set-Path $nfo
				$typeInfo = Get-TVrage-Data $nfo['url']
			}
			# Will now unpack to $path
			if($UNPACK -eq "YES"){
				Extract-Rar $path
			}
			# And create symlinks
			Create-SymLinks $path $typeInfo
		}
	}
}