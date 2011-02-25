<#
.SYNOPSIS
Automates the logic of sorting downloaded torrents.
.DESCRIPTION
Using this script with your torrent client allows it to automatically sort the torrents. 
The scripts can sort mulitple types and you can add types if needed.

ADD THIS in uTorrent

> powershell.exe C:\path\to\sort.ps1 -DIR '%D'
---------------------------------------------------------------------------

To get more information about the PARAMS

> powershell get-help .\PathTOScript.ps1 -detailed
--------------------------------------------------
#>

Param(
	# This is the DirName of the torrentParameter sent from the Client
	[parameter(Mandatory = $true)][String]$DIR,
	# Want to test for nukes? Will add nuke reason (if any) to symlink
	$TEST_NUKE = $true,
	# If the rls is a nuke, remove it and exit?
	$REMOVE_NUKE = $false,
	# Want to sort by Movie/DVDR/by_genre, /Movie/Xvid/by_genre ?
	$SORT_BY_RLS_TYPE = $false,
	# Want to automatically unpack the rls?
	$UNPACK = $false,
	# Set the basedir for sorting
	$BASE_DIR = "C:\media",
	# Specify the winrar path
	$WINRAR = "C:\Program Files\WinRAR\",
	# Set prefix for sorting folder
	$SORT_PREFIX = "by_"
)

# Script configs
$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition;
# Add paths to env:path
$env:Path = $env:Path + ";$scriptPath\bin"
$env:Path = $env:Path + ";$WINRAR"
# Scene standards references: http://scenerules.irc.gs/

# Extract data from nfo
function Extract-Nfo-Data{
	param($extractParameters)
	foreach($nfoItem in $nfoItems){
		$extractParameters.GetEnumerator() | Foreach-Object {
		if($_.Value){
			foreach($extract in $_.Value){
				# Extract parameter
				$extractThis = $nfoItem | Select-String  ([regex] "(?i:(?:$extract)\W+([\w\-,.//\.]+(?:\s[\w\-,\.]+)*))") | select -exp Matches
				$extractThis = $extractThis.Value -replace ([regex]"(?i:(?:$extract[\s\/]*\($extract\)|$extract))") -replace ([regex]'([^a-zA-z0-9\/\,\s//*])') -replace "\[|]" -replace "/", "," 
				foreach($item in $extractThis){ 
					$item = $item.trimStart().trimEnd().toLower()
					$item = (get-culture).TextInfo.ToTitleCase($item)
					$item = $item -replace ", ", "," -replace " ", "_"
					if($item){
						$item = $item.split(",")
						if($item){	
							$data += @{$_.Key = @($item)}
						}
					}	
				}
			}
		}
	}
}
return $data
}
# Unrars every *.rar file recursivly. that is, CD1 CD2, subs and so on
function Extract-Rar{
	
	if(-not(Test-Path $unpackDIR)){
			md $unpackDIR;
	}
	$rarItems = Get-ChildItem $DIR -recurse -Filter *.rar | Select-Object -ExpandProperty FullName
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
# Formats output for cmd box
function Format-Output{
	param($delim,$output)
	if($output.length -gt 78){
		$output = $output.substring(0, 78)
	}
	$count = [Math]::Truncate((80-$output.length)/2)
	return $("$delim"*$count)+$output+"$delim"*($count-1)
}
# Handle the nuke rls
function Handle-Nuked-Rls{
	Format-Output "#" " $rlsName is NUKED!"
	Format-Output " " "REASON:"
	Format-Output  " " "$nukedMsg"
	if($REMOVE_NUKE){
		Format-Output  " " "Attempting to remove the torrent+data"
		rm $rlsDIR* #-recurse
		if(Test-Path $rlsDIR){
			Format-Output " " "Could not remove the data in"
			Format-Output " " "$rlsDIR"
		}else{
			Format-Output " " "Removed all data in $rlsDIR"
		}
		# Exit here, no need to to futher investigations
		Format-Output "#" "End of NUKE message"
		break
	}
	Format-Output "#" " End of NUKE message "
}
# Set the Paths for every type
function Set-Paths{
	param($rlsName, $type, $data)
	$base = @{	MOVIE="$BASE_DIR\movie" 
				DVDR="$BASE_DIR\movie\dvdr"
				x264="$BASE_DIR\movie\x264"
				BluRayRip="$BASE_DIR\movie\blurayrip"
				BluRay="$BASE_DIR\movie\bluray"
				Xvid="$BASE_DIR\movie\xvid"
				TV="$BASE_DIR\tv" 
				MUSIC="$BASE_DIR\music"
				MVID="$BASE_DIR\mvid"
				MDVDR="$BASE_DIR\mdvdr"
				HDTV="$BASE_DIR\tv\hdtv\"
				UNKNOWN="$BASE_DIR\$type\unknown"
				UNPACK="$BASE_DIR\$type\unpacked"
				XXX="$BASE_DIR\xxx"
				GAME="$BASE_DIR\games"
				}
	if($data -AND ($type)){ 
		$data.GetEnumerator() | Foreach-Object {
			if($_.Value){
				foreach($val in $_.Value){
					$path = $SORT_PREFIX+$_.Key+"\"+$val -replace " ", "_"
					$path = join-path $base.$type $path\$rlsName
					$paths += ,($path)
				}
			}
		}
	}else{
		$paths = join-path $base.Unknown $rlsName
	}
	return $paths
}
# Get NFO
function Get-NFO{
	if(Test-Path $DIR){
		# Find an nfo in the DIR(s)
		$nfoItems = get-childitem $DIR -recurse | where {$_.extension -eq ".nfo"} 
		if($nfoItems){
			return $nfoItems
		}
	}
}    
# Determine tv source for data
function Get-Tv-Data{
	param($url)
	if($url -match "tvrage.com/shows/id-"){
		$url = $url -replace ([regex]".*//.*/id-")
		return $data = (Get-TvRage-XML-Data $url)
	}elseif($url -match "tvrage.com/"){
		$url = $url -replace ([regex]".*//.*/")
		return $data = (Get-TvRage-QuickData $url)
	}
}
# Get XML data from TvRage
function Get-TvRage-XML-Data{
	param($url)
	$id = ([regex]'([0-9]+(?:[0-9]*)?)').Match($url).value.trim()
	$url = "http://services.tvrage.com/feeds/showinfo.php?sid=$id"
	$doc = New-Object System.Xml.XmlDocument
	$doc.Load($url)
	$title = $doc.Showinfo.showname
	$genres = $doc.Showinfo.genres.genre  -replace '/', "&"
	$group = ([regex]'([a-zA-Z]*)$').Match($DIR).value.trim()
	$season = ([regex]'(\.S\d*?E\d*?\.)|(\.S\d*?\.DVDRip)|(\.S\d*?D\d*?\.)|(\.E\d*\.)|(\.S\d*?\.Disc\d*?\.)|(\.\d*?x\d*?\.)|(\.PDTV|(\.HDTV\.))').Match($DIR).value.trim()
	$season = ([regex]'(S\d*)|(\d*x)').Match($season).value.trim()
	$season = ([regex]'([0-9]+(?:[0-9]*)?)').Match($season).value.trim()
	return $ret = @{
					title = "$title\season_$season"
					genre = $genres
					group = $group
					} 
}
# Create symlinks for the torrent, this will link to the unpacked dir
function Create-SymLinks{
	Param($data)
	if($data){
		foreach ($key in $data) {
			if(-not (Test-Path $key)){
				$rlsName = $key -replace ".*\\.*\\"
				$newDir = $key -replace $rlsName
				md $newDir  -Force -ErrorAction Stop
				mklnk -a $DIR c:\windows\explorer.exe $key
			}
		}
	}
}
# Get data from IMDB, based on nfo URL
function Get-IMDb-Data{ 
    param([string] $url)
	# using akas instead of www. To get orignial titles
	$url = $url -replace "www", "akas"
	$data = (get-webfile $url -passthru )
	# Extract data
	$title = ([regex]'(?<=<h1 class="header">)([\S\s]*?)(?=<span>)').Match($data).value.trim() -replace "&#x27;" -replace ":"
	$year = ([regex]'(?<=<span>[(]<a href="/year/)(([\S\s]*?)+)(?=/">)').Match($data).value.trim().TrimEnd()
	$rating = ([regex] '(?<=<span class="rating-rating">)(([\S\s]*?)+)(?=<span>)').Match($data).value.trim()
	$votes = ([regex] '([0-9]+(?:\,[0-9]*)?)(?:\svotes)').Match($data).value.trim() -replace " votes"
	$genre = ([regex]'(?<=<a href="/genre/)(([\S\s]*?)+)(?=")').matches($data)
	$metascore = ([regex]'(?<=<span class="nobr">Metascore\S\s*<strong>)([\S\s]*?)(?=</strong>)').Match($data).value.trim()
	$byTitle = $title.substring(0,1)
	$group = ([regex]'([a-zA-Z]*)$').Match($DIR).value.trim()

	return $ret = @{
					title = "$byTitle\$title"
					year = $year
					rating = $rating
					votes = $votes
					metascore = $metascore
					genre = $genre
					group = $group
					}

}
# Get data from tvRage, based on nfo URL
function Get-TvRage-QuickData { 
	Param([String]$url)
	$data = (get-webfile http://services.tvrage.com/tools/quickinfo.php?show=$url -passthru )
	$genres = $data | Select-String  ([regex]"(?<=Genres\@)(?<text>.*)") | select -exp Matches | select -exp Value	
	$genres = $genres.split("|")|%{$_.trim() }
	$title = $data | Select-String  ([regex]"(?<=Show Name\@)(?<text>.*)") | select -exp Matches | select -exp Value
	$group = ([regex]'([a-zA-Z]*)$').Match($DIR).value.trim();
	$season = ([regex]'(\.S\d*?E\d*?\.)|(\.S\d*?\.DVDRip)|(\.S\d*?D\d*?\.)|(\.E\d*\.)|(\.S\d*?\.Disc\d*?\.)|(\.\d*?x\d*?\.)|(\.PDTV|(\.HDTV\.))').Match($DIR).value.trim()
	$season = ([regex]'(S\d*)|(\d*x)').Match($season).value.trim()
	$season = ([regex]'([0-9]+(?:[0-9]*)?)').Match($season).value.trim()
	return $ret = @{
					title = "$title\season_$season"
					genre = $genres
					group = $group
	} 
}
# Test for Nuke
function Test-Nuke{
	<#
	Test if the Rls is nuked, check it by an preDB, here Iv choosen an public one.
	You can change the url or even change it to an irc. Whatever you like
	#>
	# Remove everything before the the last \ and change the - to a + (we dont want to exclude)
	$rlsName = $DIR -replace '.*\\.*\\' -replace '-', '+'
	$nukeUrl = "http://www.orlydb.com/?q="
	<#
		We are are dealing with standards here, and urlencode should not be needed
		# Load System Web
		# $null = [Reflection.Assembly]::LoadWithPartialName("System.Web")
		# $rlsName = [System.Web.HttpUtility]::UrlEncode($rlsName)
	#>
	
	$nukeWebFile = (get-webfile $nukeUrl+$rlsName -passthru )
	$nukeReason = ([regex]'(?<=<span class="nukeright"><span class="nuke">)([\S\s]*?)(?=\</span></span>)').Match($nukeWebFile).value.trim();
	if($nukeReason){
		return "$nukeReason"
	}

}
# Test for Game (game)
function Test-Game{
	param($gameArrayToTest)
	foreach($game in $gameArrayToTest){
		$isGame = $DIR | Select-String  ([regex]"(?i)$game\-") | select -exp Matches | select -exp value
		if($isGame){ return $game}
	}
}
# Test for DVDRip, Xvid (movie)
function Test-Xvid{

	return $isXvid = $DIR | Select-String  ([regex]"(?i)\.(DVDScr|DVDScreener|DVDRip)\.(Xvid)(\.AC3(\.|\-)|\-)") | select -exp Matches | select -exp value

}
# Test for XXX (movie)
function Test-XXX{

	return $isXXX = $DIR | Select-String  ([regex]"(?i)\.XXX\.") | select -exp Matches | select -exp value
}
# Test BluRay (movie)
function Test-BluRay{
<#
		Note: Only Completes!
#>
	return $isBluRay = $DIR | Select-String  ([regex]"\.(COMPLETE.BLURAY)\-") | select -exp Matches | select -exp value
	
}	
# Test BluRay Rip (movie)
function Test-BD{
<#
	Notes: 
		Includes BR rips aswell!
		BRRip = An XviD encode from a Blu-Ray release (i.e. a 1080p *.mkv file).
		BDRip = An XviD encode directly from a source Blu-Ray disk
		They should not be confused with genuine Blu-Ray rips in 1080p, which are usually done in native Blu-Ray files, or as H.264 *.mkv files.		
#>
	return $isBD = $DIR | Select-String  ([regex]"(720p|480p)?\.(BDRip|BRRip)\.(Xvid)(\.AC3\-|\-)") | select -exp Matches | select -exp value

}
# Test x264 (movie)
function Test-x264{
<#
	Use with care, can match TV aswell
	Note: Includes BR rips aswell!
#>
	return $isx264 = $DIR | Select-String  ([regex]"(720p|1080p)\.(BluRay|BDRip|BRRip)\.(x264)\-") | select -exp Matches | select -exp value
	
}
# Test for MVID (MusicVideo)
function Test-MVID{
<#	
	Scene directives
	Directory Naming: 
	  - The following are examples of the REQUIRED naming scheme:
		  ArtistName-SongTitle-FORMAT-YEAR-GROUP
		  ArtistName-SongTitle-HDTV-FORMAT-YEAR-GROUP
		  ArtistName-SongTitle-DVDRIP-FORMAT-YEAR-GROUP
		  ArtistName-SongTitle_(Tonight_Show_01-01-07)-FORMAT-YEAR-GROUP
		  ArtistName-SongTitle_(49th_Annual_Grammys)-FORMAT-YEAR-GROUP
		  ArtistName-SongTitle_(2007_MTV_Music_Awards)-FORMAT-YEAR-GROUP
		  ArtistName-SongTitle-PROPER-FORMAT-YEAR-GROUP
		  ArtistName-SongTitle-BOOTLEG-FORMAT-YEAR-GROUP
	My Notes:
		- x264
			Allthough the standards says x264 is the only allowed codecs now days, there can be some old rls using xvid or svcd
#>
	return $isMVID = $DIR | Select-String  ([regex]"([\W\w]*)-([\W\w]*)-(x264|xvid|svcd)-(\w*\d*)") | select -exp Matches | select -exp value
	
}
# Test TV
function Test-TV{

	<#
	Scene directives
	Directory Naming:
		Numbering for episodic programming shall be in the format S00E00 or 0x00 
		(or S00E0000x000 for series with more than 99 episodes per season)                                                               
		- Number for date-based programming including sports and variety shows shall be in the format YYYY.MM.DD                                     
		  Other numbering formats such as S00.E00, DD.MM.YY, and MM.DD.YY are not permitted  
	#>
	# Check by season ([Ss]\d*[EeDd]\d*)|(\d?\d[x]\d?)
	# Diffrent types of tv makes it harder. (\.S\d*?E\d*?\.)|(\.S\d*?D\d*?\.)|(\.E\d*\.)|(\.S\d*?\.Disc\d*?\.)|(\.\d*?x\d*?\.)|(\.PDTV|(\.HDTV\.))
	if($isTv = $nfoItem | Select-String "(tvrage.com/shows/id\-([\S]*))" | select -exp Matches | select -exp value){
		return $isTV
	}
	elseif($isTv = $nfoItem | Select-String "(tvrage.com/([\S]*))" | select -exp Matches | select -exp value){
		return $isTV
	}
	elseif($isTV = $DIR | Select-String  ([regex]"(\.S\d*E\d*\.)|(\.S\d*E\d*E\d*\.)|(\.S\d*?\.DVDRip)|(\.S\d*D\d*\.)|(\.E\d*\.)|(\.S\d*\.Disc\d*\.)|(\.\d*x\d*\.).*") | select -exp Matches | select -exp value){
		# hm, the dir was a tv match, but no nice tvrage url. Lets make one
		$TVs = $rlsName | Select-String ([regex]"(.*?)(?:\$isTv)") | select -exp Matches | select -exp value
		$isTV = $TVs -replace $isTV
		return "tvrage.com/$isTV"
	}elseif($isTV = $DIR | Select-String  ([regex]"(\.|.\d.*\.)(PDTV|HDTV).*") | select -exp Matches | select -exp value){
		return "Unknown"
	}	
}
# Test movie
function Test-Movie{

<#
	How to properly check if its an movie? The only sane method would be to find an imdb link...
	Sometimes TV rips use imdb links, so we have to read the content in and exclude TV.
	But, to limit resources, lets check that it doesnt match tv regex first.
#>	
	
	return $isMovie = $nfoItem | Select-String "(http(s)?://)?(www.)?imdb.com/([\S]*)" | select -exp Matches | select -exp value
}
# Test for an mp3 file
function Test-MP3{
	if(Test-Path $DIR){
		$mp3Items = Get-ChildItem $DIR -Recurse -Filter *.mp3 | Group-Object DirectoryName | ForEach-Object { $_.Group | Select-Object DirectoryName, FullName -First 1}
		if($mp3Items){
			return $mp3Items
		}
	}
}
# Get MP3 Label
function Get-MP3Label{
	foreach($nfoItem in $nfoItems){
		# Extract Label
		$labels = $nfoItem | Select-String  ([regex]"(?i:(?:company[\s\/]*\(label\)|label|lable|company|Record co\s*|l a b e l|Catalog Number)\W+([\w\'-.//]+(?:\s[\w\'-]+)*))") | select -exp Matches
		$labels = $labels.Value -replace ([regex]'(?i:(?:company[\s\/]*\(label\)|label|lable|company|Record co\s*|l a b e l|Catalog Number))') -replace ([regex]'([^a-zA-z0-9\s//*])') -replace "\[|]" -replace "n/a" -replace "na" -replace "none" -replace "genre"		
		$labels = $labels.split("/")
		foreach($label in $labels){ 
			$label = $label.trimStart().trimEnd().toLower()
			$label = (get-culture).TextInfo.ToTitleCase($label)
			$label = $label -replace " ", "_"
			if($label){
				$labled = @{label= $label}	
			}
		}	
	}
return $labled
}
# Add the symlink to the mp3 archive
function Get-MP3-Data{
	Param($mp3Items)
	# Load the id3 lib	
	$null = [Reflection.Assembly]::LoadFrom( "$scriptPath\bin\taglib-sharp.dll") 
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
				$archiveToDirs = @{
									genre = $genre
									year = $year
									artist = "$byArtist\$artist"
									group = $group
								   }
				return $archiveToDirs
			}
		}
	}
}
# Test for MDVDR (MusicDVD)
function Test-MDVDR{
<#
	Scene directives
	Directory Naming:
	  - Complete DVD-5 releases should be named as:
          MOVIE.TITLE.YEAR.REGION.FULL.MDVDR-GROUP or 
          MOVIE.TITLE.YEAR.REGION.COMPLETE.MDVDR-GROUP
      - Non-complete DVD-5 releases should be named as:
          MOVIE.TITLE.YEAR.REGION.MDVDR-GROUP
	  - Untouched DVD-9 releases should be named either:
          MOVIE.TITLE.YEAR.REGION.MDVD9-GROUP or 
          MOVIE.TITLE.YEAR.REGION.DVD9.MDVDR-GROUP
#>
	return $isMDVDR = $DIR | Select-String  ([regex]"(\.MDVD9\-|\.MDVDR\-)") | select -exp Matches | select -exp value

}
# Test DVDR
function Test-DVDR{
<#	
	Scene directives
	Directory Naming:
	  - Directory name MUST include video standard (NTSC or PAL) except for first release of a title in regards to a retail release.
	  - Releases are to be named as: MOVIE.TITLE.YEAR.STANDARD.DVDR-GROUP
#>
	return $isDVDR = $DIR | Select-String  ([regex]"(\.DVDR|DVD9\-)") | select -exp Matches | select -exp value


}
# Test for 720p x264 HDTV
function Test-HDTV{
<#
 Scene directives
 Directory Naming: 
   - Show.Name.SXXEXX.720p.HDTV.x264-GROUP for normal series 
   - Show.Name.YYYY-MM-DD.720p.HDTV.x264-GROUP for sports/daily shows
   - Show.Name.PartXX.720p.HDTV.x264-GROUP for miniseries       
#>
	return $isHDTV = $DIR | Select-String  ([regex]"(720p).(HDTV).(x264)") | select -exp Matches | select -exp value
}
# Get-WebFile (aka wget for PowerShell)
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
         "{0} (PowerShell {1}; .NET CLR {2}; {3}; IMDB DATA)" -f $UserAgent, 
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
#################################
# Test the rls and determine type
function Test-Rls{
	Param($nfoItems=(Get-NFO))
	foreach($nfoItem in $nfoItems){
		$rlsDIR = $DIR
		$rlsName = $DIR -replace '.*\\.*\\'
		if($TEST_NUKE -AND ($nukedMsg = Test-Nuke)){
			Handle-Nuked-Rls
			$rlsName = $rlsName -replace $rlsName, "NUKED-$nukedMsg-$rlsName"
		}
		if($tv = Test-TV){
			$type = "TV"
			if($SORT_BY_RLS_TYPE -AND(Test-HDTV)){
				$type = "HDTV"
			}
			if($tv -match "Unknown"){
				$paths = Set-Paths $rlsName $type
			}else{
				$paths = Set-Paths $rlsName $type (Get-Tv-Data $tv)
			}	
		}elseif($MP3 = Test-MP3){
			$type = "MUSIC"
			$paths = Set-Paths $rlsName $type (Get-MP3-Data $MP3)
			$paths = Set-Paths $rlsName $type (Get-MP3Label)
		}elseif($movie = Test-Movie){
			$type = "MOVIE"
			if($SORT_BY_RLS_TYPE){
				if(Test-DVDR){
					$type = "DVDR"
				}elseif(Test-x264){
					$type = "x264"
				}elseif(Test-BD){
					$type = "BluRayRip"
				}elseif(Test-BluRay){
					$type = "BluRay"
				}elseif(Test-Xvid){
					$type ="Xvid"
				}
			}
			$paths = Set-Paths $rlsName $type (Get-IMDb-Data $movie)		
		}elseif(Test-MDVDR){
			$type = "MDVDR"
			$paths = Set-Paths $rlsName $type (Extract-Nfo-Data @{artist = @("artist"); genre = @("genre"); year = @("year"); label = @("label"); language = @("language")})
		}elseif(Test-MVID){
			$type = "MVID"
			$paths = Set-Paths $rlsName $type (Extract-Nfo-Data @{artist = @("artist"); genre = @("genre"); year = @("year"); label = @("label"); language = @("language")})
		}elseif(Test-XXX){
			$type = "XXX"
			$paths = Set-Paths $rlsName $type (Extract-Nfo-Data @{cast = @("cast", "featuring", "starring"); studio = @("studio", "company"); genre = @("category", "genre")})
		}elseif($game = (Test-Game  @("PS3", "NDS", "XBOX360"))){
			$type = "GAME"
			$paths = Set-Paths $rlsName $type (Extract-Nfo-Data @{platform = @("platform"); genre = @("category", "genre")})
		}
		if($UNPACK -AND(Get-ChildItem $DIR -recurse -Filter *.rar | Select-Object -ExpandProperty FullName)){ 
			
				$type = $type.toLower()
				$unpackDIR = "$BASE_DIR\$type\unpacked\$rlsName"
				(Extract-Rar)
				$DIR = $unpackDIR
			
		}	
		if(!($type)){
			Format-Output " " "The torrent is of UNKNOWN type! I will not sort!"
			break
		}
	}

	$ret = @{data = $paths
			 nuked = $nukedDir
			 type = $type 
			}
	Format-Output "#" " $type "
	(Create-SymLinks $ret.data)
	return
}
Test-Rls
