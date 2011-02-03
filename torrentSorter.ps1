#requires -version 2
Param([String]$TORRENT_NAME, [String]$TORRENT_DIR);
##Edit this please:###############################################################

	# Set the baseDir for your sorting and unpacking							
		$baseMovieDIR = "F:\media\movies\";										
		$baseTvDIR = "F:\media\tv\"												
	# set the correct path to winRar	
		$winRar = "C:\Program Files\WinRAR\";				
	# Want to enable loggin if there is errors?	
	# If Set to NO, it will only be outputted in console
		$enableLogging = "NO";
##End of edit#####################################################################

# Set scriptPath first
$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition;
# LoggFunction
function logThis(){
	Param([string] $logThis)
	$date = Get-Date;
	$loggStr = "$date : $logThis"
	if($enableLogging -eq "YES"){
		$loggStr >> $scriptPath\logFile.txt
	}else{
		write-output $logThis;
	}
}

# Test crucial dirs and drives
$driveOnMovieBase = (New-Object System.IO.DriveInfo($baseMovieDIR)).DriveType -ne 'NoRootDirectory';
$driveOnTvBase = (New-Object System.IO.DriveInfo($baseTvDIR)).DriveType -ne 'NoRootDirectory';
if(-not($driveOnMovieBase)){ logThis "No Movie Drive!";break  }
if(-not($driveOnTvBase)){ logThis "No Tv Drive!"; break; }     
if(-not(Test-Path $winRar)){ logThis "Can not find WinRar in path!"; break; }
# Test End

# Add paths to env:path
	$env:Path = $env:Path + ";$scriptPath\bin";		
	$env:Path = $env:Path + ";$winRar";	

# Set dirs, create if not exist
# Note: You can edit the names of these, but that would break the script. 
# To edit or add, you need to edit/add regex in getData
$unpackMovieDIR = $baseMovieDIR+"unpacked\"+$TORRENT_NAME;
$unpackTvDIR = $baseTvDIR+"unpacked\"+$TORRENT_NAME;
$movieDIRS = ("unpacked", "sorted\genre", "sorted\rating", "sorted\year", "sorted\title", "sorted\metascore");
$tvDIRS = ("unpacked", "sorted\genre", "sorted\title");
# Create the basedirs
if(-not(Test-Path $baseMovieDIR)){ 
	mkdir $baseMovieDIR;
	if(-not(Test-Path $baseMovieDIR)){ 
		logThis "Could not create Movie BaseDir!" 
		break;
	}
}
# Foreach dir in movieDIRS, create them if not exist
foreach($DIR in $movieDIRS){
	$symMovieDIR = $baseMovieDIR+$DIR;
	
	if(-not (Test-Path $symMovieDIR)){
		mkdir $symMovieDIR;
		if(-not (Test-Path $symMovieDIR)){
			break {logThis "Could not create Movie DIRS!"}
		}
	}
}
# Foreach dir in tvDIRS, create them if not exist
foreach($DIR in $tvDIRS){
	$symTvDIR = $baseTvDir+$DIR;
	
	if(-not (Test-Path $symTvDIR)){
		mkdir $symTvDIR;
		if(-not (Test-Path $symTvDIR)){
			break {logThis "Could not create Tv DIRS!"}
		}
	}
}

# check the NFO for imdb och tvrage link
function checkNFO(){
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
function setPath(){
	Param($res)
	if($res['type'] -eq "Tv"){
		return $unpackTvDIR;
	}
	if($res['type'] -eq "Movie"){
		return $unpackMovieDIR;
	}
	if(!$res['type']){
		return "Error";
	}
}

# Unrar Function. 
# Unrars every *.rar file recursivly. that is, CD1 CD2, subs and so on
function doUnrar(){
	Param($unpackDir)
	if(Test-Path $TORRENT_DIR){
		if(-not($unpackDir -eq "Error")){
			
			if(-not(Test-Path $unpackDIR)){
				md $unpackDIR;
			}
			$rarItems = Get-ChildItem $TORRENT_DIR -recurse -Filter *.rar | Select-Object -ExpandProperty FullName
			foreach($item in $rarItems){
				& unrar e -o- $item $unpackDIR;
			}
			# Unrars sub item and then removes the package. This is in the unpacked folder and not Original folder
			$subItems = Get-ChildItem $unpackDIR -recurse -Filter *.rar | Select-Object -ExpandProperty FullName
			if($subItems){
				foreach($item in $subItems){
					& unrar e -o- $item $unpackDIR;
					ri $item;
				}
			}
		}
	}
}

# Get data from IMDB, based on nfo URL
function getIMDbData { 
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
					bytitle = $baseMovieDIR+"sorted\title\"+$byTitle+"\"+$title;
					year = $baseMovieDIR+"sorted\year\"+$year;
					rating = $baseMovieDIR+"sorted\rating\"+$rating; 
					metascore = $baseMovieDIR+"sorted\metascore\"+$metascore;
					genre1 = $baseMovieDIR+"sorted\genre\"+$genre1;
					genre2 = $baseMovieDIR+"sorted\genre\"+$genre2;
					genre3 = $baseMovieDIR+"sorted\genre\"+$genre3}; 

	}

# Get data from tvRage, based on nfo URL
function getTvRageData { 
	Param($url)
	$season = $TORRENT_NAME -replace '.*s(.*)e.*','$1';
	$title, $genres = (get-webfile http://services.tvrage.com/tools/quickinfo.php?show=$url -passthru ) -split "`n" | select -index 1,13
	$title = $title -replace "Show Name@", "";
	$title = $title -replace ":", "";
	$title = $title -replace " ", "_";
	$genres = $genres -replace "Genres@", "";
	$genre1,$genre2,$genre3 = $genres.split("|");
	return $ret = @{title = $title;
					season = $baseTvDIR+"sorted\title\"+$title+"\season_"+$season;
					genre1 = $baseTvDIR+"sorted\genre\"+$genre1.trim();
					genre2 = $baseTvDIR+"sorted\genre\"+$genre2.trim();
					genre3 = $baseTvDIR+"sorted\genre\"+$genre3.trim()}; 
	}

# Create symlinks for the torrent, this will link to the unpacked dir	
function createSymLinks(){
	Param($unpackDir, $data);
	if($data){
		foreach ($key in $data.keys -ne 'title') {
			if(-not (Test-Path $data.$key)){
				mkdir $data.$key; 
			}
			$value = $data.$key;
			$dir = $value+"\"+$TORRENT_NAME;
			mklnk -a $unpackDir c:\windows\explorer.exe $dir;
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
$nfo = checkNFO;
if(!$nfo){

	logThis "Could not find an NFO!";

}elseif(!$nfo['type']){
	logThis "Could not determine type!";
}elseif(!$nfo['url']){
	logThis "Could not find an URL!";
}elseif($nfo){
	if(!$nfo['type']){
		logThis "Could not determine type!";
	}elseif(!$nfo['url']){
		logThis "Could not find an URL!";
		
	}else{
	
		# First, check the type from nfo and set path
		if($nfo['type'] -eq "Movie"){
			$path = setPath $nfo;
			$typeInfo = getIMDbData $nfo['url'];
		}elseif($nfo['type'] -eq "Tv"){
			$path = setPath $nfo;
			$typeInfo = getTvRageData $nfo['url'];
		}
		# Will now unpack to $path
		doUnrar $path;
		# And create symlinks
		createSymLinks $path $typeInfo;
	}
}