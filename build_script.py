"""Build a self-contained, plain AppleScript launcher for Script Editor."""
from pathlib import Path
import subprocess

root = Path(__file__).resolve().parent
core = (root / 'ExtractKeynoteTextCore.applescript').read_text()
literal = '"' + core.replace('\\', '\\\\').replace('"', '\\"') + '"'
launcher = '''use scripting additions

-- 選檔介面不載入 Foundation / AppKit；核心由獨立 osascript 執行。
-- 核心已內嵌，單獨使用這個檔案即可，不需另外安裝任何套件。
on run argv
	set workDir to ""
	set fileHandle to missing value
	try
		set argumentCount to count of argv
	on error
		set argv to {}
		set argumentCount to 0
	end try
	try
		if argumentCount = 0 then
			set sourceFile to choose file with prompt "選擇要擷取文字的 Keynote 或 PowerPoint 簡報" of type {"key", "pptx"}
			set sourcePath to POSIX path of sourceFile
			set imageChoice to button returned of (display dialog "是否辨識投影片中的圖片文字？" buttons {"取消", "跳過圖片", "辨識圖片"} default button "辨識圖片" cancel button "取消")
			set sourceName to name of (info for sourceFile)
			if sourceName ends with ".key" then set sourceName to text 1 thru -5 of sourceName
			if sourceName ends with ".pptx" then set sourceName to text 1 thru -6 of sourceName
			set sourceFolder to do shell script "/usr/bin/dirname " & quoted form of sourcePath
			set outputFile to choose file name with prompt "儲存擷取結果（請使用新檔名）" default name (sourceName & "_文字.txt") default location ((POSIX file sourceFolder) as alias)
			set workerArgs to {sourcePath, POSIX path of outputFile}
			if imageChoice is "跳過圖片" then set end of workerArgs to "--skip-images"
		else
			if argumentCount > 3 then error "用法：osascript ExtractKeynoteText.applescript 輸入.key或.pptx [輸出.txt] [--skip-images]"
			set workerArgs to argv
		end if
		set workDir to do shell script "/usr/bin/mktemp -d /tmp/keynote-launcher.XXXXXXXX"
		set workerPath to workDir & "/extract.applescript"
		set fileHandle to open for access (POSIX file workerPath) with write permission
		write my workerSource to fileHandle as «class utf8»
		close access fileHandle
		set fileHandle to missing value
		set commandText to "/usr/bin/osascript " & quoted form of workerPath
		repeat with argumentText in workerArgs
			set commandText to commandText & " " & quoted form of (argumentText as text)
		end repeat
		set outputPath to do shell script commandText
		my cleanupLauncher(workDir)
		if argumentCount = 0 then display dialog "擷取完成。" & linefeed & outputPath buttons {"好"} default button 1
		return outputPath
	on error errorMessage number errorNumber
		if fileHandle is not missing value then
			try
				close access fileHandle
			end try
		end if
		my cleanupLauncher(workDir)
		if errorNumber = -128 then return
		error errorMessage number errorNumber
	end try
end run

on cleanupLauncher(workDir)
	if workDir starts with "/tmp/keynote-launcher." then
		do shell script "/bin/rm -rf " & quoted form of workDir
	end if
end cleanupLauncher

-- 由 build_script.py 從 ExtractKeynoteTextCore.applescript 產生。
property workerSource : ''' + literal + '\n'
source = root / 'ExtractKeynoteText.applescript'
source.write_text(launcher)
subprocess.run(['/usr/bin/osacompile', '-o', str(root / 'ExtractKeynoteText.scpt'), str(source)], check=True)
