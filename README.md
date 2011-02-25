SYNOPSIS 
========

The (u)TorrentSorter Automates the logic of sorting downloaded torrents.
### What it sorts
It sorts practically every known type like movies, music, games, tv shows, mvid, mdvdr, xxx and so on, by title, genre, year, group, cast, studio, labels etc.
### Can i add types?
You can easily add new types in the script.  
### Nukes?
If enabled the script can check if the rls is a nuke, then handle that by removing it or renaming it with the nuke reason.
USAGE
======

###ADD THIS in your torrentclient	

	> powershell.exe C:\path\to\torrentSorter.ps1 -DIR '%D'

###To get more information about the PARAMS

	> powershell get-help .\path\to\torrentSorter.ps1 -detailed


DEPENDENCIES
============

	bin\mklnk.exe (included)
	
	bin\taglib-sharp.dll(included)
	
	winRAR (not included)
	
	uTorrent (not included) 	
	
###Authors:

	mklnk is Copyright (c) 2005-2006 Ross Smith II (http://smithii.com) All Rights Reserved
	taglib-sharp.dll is Copyright (c) http://code.google.com/p/thelastripper


INFORMATION
===========
### Implementation
So, almost everybody run Windows 7 or Vista these days. If you do, Powershell is inlcuded and to test that
you have it installed, open a command promt run -> cmd) and type 

	powershell $psversiontable

####Output:

	C:\Windows\System32>powershell $psversiontable

	Name                           Value
	----                           -----
	PSVersion                      2.0


So, if this works, and PSVersion = 2.0 we are good to go.

###Then:

	Type> Get-ExecutionPolicy
	Output> Restricted

If the output says Restricted, it means that the default value is still set. Change this by doing

	Type> Set-ExecutionPolicy RemoteSigned
Read more here: http://technet.microsoft.com/en-us/library/ee176949.aspx

###Else:

####Goto 
	http://support.microsoft.com/kb/968929
####scroll to
	Windows Management Framework Core (WinRM 2.0 and Windows PowerShell 2.0)
Download for your windows version
####Install
###OR: 
Do a system update. It should be included.

LINKS
====
###Github:
	https://github.com/mline/uTorrentSorter