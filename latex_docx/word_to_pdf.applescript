on run argv
	set inPath to item 1 of argv
	set outPath to item 2 of argv
	tell application "Microsoft Word"
		activate
		open POSIX file inPath
		set theDoc to active document
		save as theDoc file name (POSIX file outPath) file format format PDF
		close theDoc saving no
	end tell
end run
