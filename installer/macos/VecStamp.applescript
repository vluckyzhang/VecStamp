-- VecStamp (Shi Yin) - macOS helper for the PowerPoint add-in.
-- PowerPoint for Mac runs VBA in a sandbox; the add-in calls these handlers
-- through AppleScriptTask for save/folder dialogs, moving exported files out of
-- the sandbox, gzip (EMZ), zip (VSDX), the colour picker, revealing files,
-- optional SVG conversion with Inkscape or LibreOffice, the "exporting..."
-- progress dialog and copying text to the clipboard.
--
-- Installed (compiled with osacompile) to:
--   ~/Library/Application Scripts/com.microsoft.Powerpoint/VecStamp.scpt
--
-- Copyright (c) 2026 vluckyzhang <vluckyzhang@gmail.com>
-- Released under the MIT License.

property sepText : "|*|"

on splitParam(theText)
	set oldDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to sepText
	set parts to text items of theText
	set AppleScript's text item delimiters to oldDelims
	return parts
end splitParam

on partOf(parts, n)
	if (count of parts) >= n then return item n of parts
	return ""
end partOf

on helperVersion(paramString)
	return "1.2.1"
end helperVersion

-- "prompt|*|default name|*|default folder" -> POSIX path, or "" when cancelled
on chooseSaveFile(paramString)
	set parts to splitParam(paramString)
	set promptText to partOf(parts, 1)
	set defName to partOf(parts, 2)
	set defFolder to partOf(parts, 3)
	try
		if defFolder is not "" then
			set folderAlias to (POSIX file defFolder) as alias
			set f to choose file name with prompt promptText default name defName default location folderAlias
		else
			set f to choose file name with prompt promptText default name defName
		end if
		return POSIX path of f
	on error errMsg number errNum
		if errNum is -128 then return ""
		try
			set f to choose file name with prompt promptText default name defName
			return POSIX path of f
		on error
			return ""
		end try
	end try
end chooseSaveFile

-- "prompt|*|default folder" -> POSIX path of the chosen folder, or ""
on chooseFolder(paramString)
	set parts to splitParam(paramString)
	set promptText to partOf(parts, 1)
	set defFolder to partOf(parts, 2)
	try
		if defFolder is not "" then
			set folderAlias to (POSIX file defFolder) as alias
			set f to choose folder with prompt promptText default location folderAlias
		else
			set f to choose folder with prompt promptText
		end if
		return POSIX path of f
	on error errMsg number errNum
		if errNum is -128 then return ""
		try
			set f to choose folder with prompt promptText
			return POSIX path of f
		on error
			return ""
		end try
	end try
end chooseFolder

-- "source|*|destination" -> "ok" or an error message
on moveFile(paramString)
	set parts to splitParam(paramString)
	try
		do shell script "/bin/mv -f " & quoted form of partOf(parts, 1) & " " & quoted form of partOf(parts, 2)
		return "ok"
	on error errMsg
		return errMsg
	end try
end moveFile

-- "source.emf|*|destination.emz" -> "ok"; the source is removed
on gzipFile(paramString)
	set parts to splitParam(paramString)
	set src to partOf(parts, 1)
	set dst to partOf(parts, 2)
	try
		do shell script "/usr/bin/gzip -c -9 " & quoted form of src & " > " & quoted form of dst & " && /bin/rm -f " & quoted form of src
		return "ok"
	on error errMsg
		return errMsg
	end try
end gzipFile

on pathExists(p)
	try
		do shell script "/bin/test -e " & quoted form of p
		return "1"
	on error
		return "0"
	end try
end pathExists

on isFolder(p)
	try
		do shell script "/bin/test -d " & quoted form of p
		return "1"
	on error
		return "0"
	end try
end isFolder

on desktopFolder(paramString)
	return POSIX path of (path to desktop folder)
end desktopFolder

on revealFile(p)
	try
		do shell script "/usr/bin/open -R " & quoted form of p
	end try
	return "ok"
end revealFile

-- opens a folder in Finder or a mailto: link in the mail app
on openPath(p)
	try
		do shell script "/usr/bin/open " & quoted form of p
	end try
	return "ok"
end openPath

-- "source.emf|*|destination.svg" -> "ok" / "missing" / error text
on inkscapeSvg(paramString)
	set parts to splitParam(paramString)
	set src to partOf(parts, 1)
	set dst to partOf(parts, 2)
	set ink to "/Applications/Inkscape.app/Contents/MacOS/inkscape"
	try
		do shell script "/bin/test -x " & quoted form of ink
	on error
		return "missing"
	end try
	try
		do shell script quoted form of ink & " " & quoted form of src & " --export-type=svg " & quoted form of ("--export-filename=" & dst)
		do shell script "/bin/test -s " & quoted form of dst
		return "ok"
	on error errMsg
		return errMsg
	end try
end inkscapeSvg

-- "source.emf|*|destination.svg" -> "ok" / "missing" / error text
on libreofficeSvg(paramString)
	set parts to splitParam(paramString)
	set src to partOf(parts, 1)
	set dst to partOf(parts, 2)
	set lo to "/Applications/LibreOffice.app/Contents/MacOS/soffice"
	try
		do shell script "/bin/test -x " & quoted form of lo
	on error
		return "missing"
	end try
	set outDir to do shell script "/usr/bin/mktemp -d /tmp/vecstamp_lo.XXXXXX"
	try
		do shell script quoted form of lo & " -env:UserInstallation=file:///tmp/vecstamp_lo_profile --headless --norestore --convert-to svg --outdir " & quoted form of outDir & " " & quoted form of src
		set baseName to do shell script "/usr/bin/basename " & quoted form of src & " .emf"
		do shell script "/bin/mv -f " & quoted form of (outDir & "/" & baseName & ".svg") & " " & quoted form of dst
		do shell script "/bin/rm -rf " & quoted form of outDir
		return "ok"
	on error errMsg
		do shell script "/bin/rm -rf " & quoted form of outDir
		return errMsg
	end try
end libreofficeSvg

-- "folder|*|package.vsdx" -> "ok"; zips an OPC package ([Content_Types].xml first)
on zipFolder(paramString)
	set parts to splitParam(paramString)
	set srcDir to partOf(parts, 1)
	set dst to partOf(parts, 2)
	try
		do shell script "/bin/rm -f " & quoted form of dst
		do shell script "cd " & quoted form of srcDir & " && /usr/bin/zip -X -q -nw " & quoted form of dst & " '[Content_Types].xml' && /usr/bin/zip -X -q -nw -D -r " & quoted form of dst & " . -x '[Content_Types].xml'"
		return "ok"
	on error errMsg
		return errMsg
	end try
end zipFolder

-- removes one of VecStamp's own temporary folders
on removePath(p)
	if p does not contain "VecStamp" then return "refused"
	try
		do shell script "/bin/rm -rf " & quoted form of p
	end try
	return "ok"
end removePath

-- "r,g,b" (0-255) -> chosen colour as "r,g,b", or "" when cancelled
on chooseColor(paramString)
	set oldDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to ","
	set parts to text items of paramString
	set AppleScript's text item delimiters to oldDelims
	set startColor to {65535, 63736, 60395}
	try
		set startColor to {(item 1 of parts as integer) * 257, (item 2 of parts as integer) * 257, (item 3 of parts as integer) * 257}
	end try
	try
		set c to choose color default color startColor
	on error
		return ""
	end try
	return ((round ((item 1 of c) / 257)) as text) & "," & ((round ((item 2 of c) / 257)) as text) & "," & ((round ((item 3 of c) / 257)) as text)
end chooseColor

-- Shows a small "please wait" dialog in a separate process; returns its process id.
-- (AppleScriptTask blocks PowerPoint while it runs, so the dialog must not belong to this script.)
on progressShow(msg)
	try
		set cmd to "/usr/bin/osascript -e 'on run argv' -e 'display dialog (item 1 of argv) with title \"VecStamp\" buttons {\"OK\"} default button 1 giving up after 600' -e 'end run' " & quoted form of msg & " >/dev/null 2>&1 & echo $!"
		return do shell script cmd
	on error
		return ""
	end try
end progressShow

on progressHide(pidText)
	try
		set n to pidText as integer
		do shell script "/bin/kill " & (n as text) & " >/dev/null 2>&1; true"
	end try
	return "ok"
end progressHide

on setClipboard(theText)
	try
		set the clipboard to theText
		return "ok"
	on error
		return ""
	end try
end setClipboard
