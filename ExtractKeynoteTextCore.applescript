use AppleScript version "2.7"
use framework "Foundation"
use framework "AppKit"
use scripting additions

-- Keynote Creator Studio 的 bundle ID。傳統 Keynote 請將本檔所有
-- com.apple.Keynote 改成 com.apple.iWork.Keynote，再重新編譯。
property shortcutName : "提取圖片文字"
property includePresenterNotes : false
property fm : missing value
property ocrCache : missing value
property warningCount : 0
property imageCount : 0
property ocrRunCount : 0

on run argv
	-- 腳本編輯器可能傳入無法解析的物件參照，而非命令列的參數清單。
	-- 先正規化，再讀取 count，讓編輯器的「執行」使用選檔模式。
	try
		if (class of argv) is not list then set argv to {}
		set argumentCount to count of argv
	on error
		set argv to {}
		set argumentCount to 0
	end try
	set processImages to true
	set exportImages to false
	set skipImages to false
	set inputArgs to {}
	repeat with argumentText in argv
		if (argumentText as text) is "--skip-images" then
			set processImages to false
			set skipImages to true
		else if (argumentText as text) is "--export-images" then
			set processImages to false
			set exportImages to true
		else
			set end of inputArgs to argumentText as text
		end if
	end repeat
	set argv to inputArgs
	set argumentCount to count of argv
	set fm to current application's NSFileManager's defaultManager()
	set ocrCache to missing value
	if processImages then set ocrCache to current application's NSMutableDictionary's dictionary()
	set warningCount to 0
	set imageCount to 0
	set ocrRunCount to 0
	set tempDir to ""
	set reportPath to ""
	set reportText to ""
	set imageFolder to ""
	set exportedImageCount to 0
	set interactiveMode to argumentCount = 0
	try
		if skipImages and exportImages then error "--skip-images 與 --export-images 不可同時使用。"
		if argumentCount > 2 then error "用法：osascript ExtractKeynoteText.applescript 輸入.key或.pptx [輸出.txt] [--skip-images | --export-images]"
		if interactiveMode then
			set sourceFile to choose file with prompt "選擇要擷取文字的 Keynote 或 PowerPoint 簡報" of type {"key", "pptx"}
			set imageChoices to choose from list {"辨識圖片", "跳過圖片", "取出圖片"} with prompt "請選擇圖片處理方式（取出圖片會存檔，不執行 OCR）：" default items {"辨識圖片"} OK button name "繼續" cancel button name "取消" multiple selections allowed false empty selection allowed false
			if imageChoices is false then return
			set processImages to (item 1 of imageChoices) is "辨識圖片"
			set exportImages to (item 1 of imageChoices) is "取出圖片"
			if processImages then set ocrCache to current application's NSMutableDictionary's dictionary()
		else
			set sourceFile to (POSIX file (my absolutePath(item 1 of argv))) as alias
		end if
		set sourcePath to POSIX path of sourceFile
		set sourceNSString to current application's NSString's stringWithString:sourcePath
		set sourceExtension to ((sourceNSString's pathExtension())'s lowercaseString()) as text
		if sourceExtension is not "key" and sourceExtension is not "pptx" then error "只支援 .key 或 .pptx 簡報：" & sourcePath
		set defaultOutput to ((sourceNSString's stringByDeletingPathExtension()) as text) & "_文字.txt"
		if interactiveMode then
			set reportPath to POSIX path of (choose file name with prompt "儲存擷取結果（請使用新檔名）" default name ((current application's NSString's stringWithString:defaultOutput)'s lastPathComponent() as text) default location ((POSIX file (sourceNSString's stringByDeletingLastPathComponent() as text)) as alias))
		else if argumentCount = 2 then
			set reportPath to my absolutePath(item 2 of argv)
		else
			set reportPath to my unusedPath(defaultOutput)
		end if
		if (fm's fileExistsAtPath:reportPath) as boolean then error "輸出檔已存在，請改用新檔名：" & reportPath
			if processImages then
				set shortcutList to paragraphs of (do shell script "/usr/bin/shortcuts list")
				if shortcutList does not contain shortcutName then error "找不到捷徑「" & shortcutName & "」。請先在捷徑 App 建立或同步此捷徑。"
			end if
			set tempDir to do shell script "/usr/bin/mktemp -d /tmp/keynote-text.XXXXXXXX"
			set pptxPath to tempDir & "/snapshot.pptx"
			set packageRoot to tempDir & "/unpacked"
		-- 只開啟與匯出，不修改、不儲存、不關閉使用者的簡報。
		with timeout of 600 seconds
			tell application id "com.apple.Keynote"
				set deck to open sourceFile
				set pageCount to count of slides of deck
					export deck to (POSIX file pptxPath) as Microsoft PowerPoint
				end tell
			end timeout
			do shell script "/usr/bin/ditto -x -k " & quoted form of pptxPath & " " & quoted form of packageRoot
			-- 必須依 presentation.xml 的關聯順序，不可用 slide*.xml 字典排序。
			set presentationPath to packageRoot & "/ppt/presentation.xml"
			set presentationXML to my readXML(presentationPath)
			set pageIDs to my nodes(presentationXML, "/*[local-name()='presentation']/*[local-name()='sldIdLst']/*[local-name()='sldId']/@r:id")
			set presentationRels to my relationships(presentationPath)
			if (pageIDs's |count|() as integer) is not pageCount then error "Keynote 與匯出 PPTX 頁數不同，已停止，避免將內容配到錯頁。"
			if processImages then
				set imageModeText to "OCR 捷徑：" & shortcutName
		else if exportImages then
			set imageFolder to my createImageFolder(reportPath)
			set imageModeText to "圖片處理：取出圖片" & linefeed & "圖片資料夾：" & imageFolder
		else
			set imageModeText to "圖片處理：已跳過"
		end if
		set reportText to "來源：" & sourcePath & linefeed & "頁數：" & pageCount & linefeed & imageModeText & linefeed & "狀態：處理中（每頁完成後更新）" & linefeed
		my writeUTF8(reportText, reportPath)
		repeat with pageIndex from 1 to pageCount
			log "擷取第 " & pageIndex & " / " & pageCount & " 頁"
			tell application id "com.apple.Keynote"
				set pageRef to slide pageIndex of deck
				set isSkipped to skipped of pageRef
			end tell
			set pageHeading to "===== 第 " & pageIndex & " 頁 ====="
			if isSkipped then set pageHeading to pageHeading & "（略過的投影片）"
			set nativeText to my containerText(pageRef)
			if nativeText is "" then set nativeText to "（無可讀取的原生文字）" & linefeed
			set pageText to linefeed & pageHeading & linefeed & linefeed & "【投影片文字】" & linefeed & nativeText
				if includePresenterNotes then
				tell application id "com.apple.Keynote" to set notesText to (get presenter notes of pageRef) as text
				if notesText is not "" then set pageText to pageText & linefeed & "【講者備忘稿】" & linefeed & notesText & linefeed
				end if
				set pageID to ((pageIDs's objectAtIndex:(pageIndex - 1))'s stringValue()) as text
				set pagePath to my relatedPath(presentationRels, pageID, presentationPath, packageRoot)
				set hyperlinkText to my pageHyperlinks(pagePath)
				if hyperlinkText is not "" then set pageText to pageText & linefeed & hyperlinkText
				if processImages or exportImages then
					set mediaPaths to my pageImages(pagePath, packageRoot)
				set imageText to ""
				repeat with imageIndex from 1 to count of mediaPaths
					set imageCount to imageCount + 1
					set mediaPath to item imageIndex of mediaPaths
					if exportImages then
						set imageName to my exportImage(mediaPath, imageFolder, pageIndex, imageIndex)
						if imageName is not "" then
							set exportedImageCount to exportedImageCount + 1
							set folderName to (current application's NSString's stringWithString:imageFolder)'s lastPathComponent() as text
							set imageText to imageText & "[圖片 " & imageIndex & "] " & folderName & "/" & imageName & linefeed
						end if
					else
						set ocrText to my recognizeImage(mediaPath, tempDir)
						if ocrText does not start with "【OCR 失敗】" and ocrText does not start with "【警告】" then
							if my hasOCRText(ocrText) then set imageText to imageText & "[圖片 " & imageIndex & "]" & linefeed & ocrText & linefeed & linefeed
						end if
					end if
				end repeat
				if imageText is not "" then set pageText to pageText & linefeed & imageText
			end if
			set reportText to reportText & pageText
			my writeUTF8(reportText, reportPath)
		end repeat
		set reportText to my replaceText(reportText, "狀態：處理中（每頁完成後更新）", "狀態：完成")
		set reportText to reportText & linefeed & "===== 擷取統計 =====" & linefeed & "投影片：" & pageCount & linefeed
		if processImages then set reportText to reportText & "各頁圖片合計（同頁相同檔案只算一次）：" & imageCount & linefeed
		if exportImages then set reportText to reportText & "已取出圖片：" & exportedImageCount & linefeed
		set reportText to reportText & "實際執行 OCR：" & ocrRunCount & linefeed & "警告：" & warningCount & linefeed
		my writeUTF8(reportText, reportPath)
		my cleanup(tempDir)
		if interactiveMode then display dialog "擷取完成。" & linefeed & "警告：" & warningCount & linefeed & reportPath buttons {"好"} default button 1
		return reportPath
	on error errorMessage number errorNumber
		if reportText is not "" then
			try
				my writeUTF8(reportText & linefeed & "【處理中止】" & errorMessage & " (" & errorNumber & ")" & linefeed, reportPath)
			end try
		end if
		my cleanup(tempDir)
		if errorNumber = -128 and interactiveMode then return
		error errorMessage number errorNumber
	end try
end run

-- 資料夾與 TXT 同層；已有同名資料夾或檔案時另取新名。
on createImageFolder(reportPath)
	set basePath to ((current application's NSString's stringWithString:reportPath)'s stringByDeletingPathExtension() as text) & "_圖片"
	set candidatePath to basePath
	set suffix to 2
	repeat while (fm's fileExistsAtPath:candidatePath) as boolean
		set candidatePath to basePath & "_" & suffix
		set suffix to suffix + 1
	end repeat
	do shell script "/bin/mkdir " & quoted form of candidatePath
	return candidatePath
end createImageFolder

on exportImage(mediaPath, imageFolder, pageIndex, imageIndex)
	if mediaPath starts with "WARNING:" then return ""
	set extensionText to (current application's NSString's stringWithString:mediaPath)'s pathExtension() as text
	set imageName to "page_" & my paddedIndex(pageIndex) & "_image_" & my paddedIndex(imageIndex)
	if extensionText is not "" then set imageName to imageName & "." & extensionText
	set {didCopy, copyError} to fm's copyItemAtPath:mediaPath toPath:(imageFolder & "/" & imageName) |error|:(reference)
	if not (didCopy as boolean) then
		set warningCount to warningCount + 1
		return ""
	end if
	return imageName
end exportImage

on paddedIndex(indexValue)
	set indexText to indexValue as text
	repeat while (count of indexText) < 3
		set indexText to "0" & indexText
	end repeat
	return indexText
end paddedIndex

-- 直接巡訪 Keynote 物件；預設標題、內文亦在 iWork items 之中，
-- 不再另外讀一次，避免重複。群組必須遞迴。
on containerText(containerRef)
	set resultText to ""
	tell application id "com.apple.Keynote" to set childItems to get every iWork item of containerRef
	repeat with childRef in childItems
		try
			tell application id "com.apple.Keynote"
				set itemClass to class of childRef
				if itemClass is group then
					set resultText to resultText & my containerText(childRef)
				else if itemClass is shape or itemClass is text item then
					set itemText to (get object text of childRef) as text
					if itemText is not "" then set resultText to resultText & itemText & linefeed & linefeed
				else if itemClass is table then
					set resultText to resultText & my tableText(childRef) & linefeed
				else if itemClass is chart then
					set resultText to resultText & "【警告】圖表文字不在 Keynote 文字介面中，需另行檢查。" & linefeed
					set my warningCount to (my warningCount) + 1
				end if
			end tell
		on error errorMessage number errorNumber
			if errorNumber = -128 then error errorMessage number errorNumber
			set warningCount to warningCount + 1
			set resultText to resultText & "【文字讀取失敗】" & errorMessage & linefeed
		end try
	end repeat
	return resultText
end containerText

on tableText(tableRef)
	set resultText to ""
	tell application id "com.apple.Keynote"
		repeat with rowRef in rows of tableRef
			set cellTexts to {}
			repeat with cellRef in cells of rowRef
				set cellText to formatted value of cellRef
				if cellText is missing value then set cellText to ""
				set end of cellTexts to cellText as text
			end repeat
			set resultText to resultText & my joinText(cellTexts, tab) & linefeed
		end repeat
	end tell
	return resultText
end tableText

-- Keynote 的 AppleScript 字典不公開文字超連結，因此從匯出的 PPTX
-- 找出含 hlinkClick 的文字 run，再由本頁 relationship 取得 URL。
on pageHyperlinks(pagePath)
	set pageXML to my readXML(pagePath)
	set pageRels to my relationships(pagePath)
	set hyperlinkNodes to my nodes(pageXML, "//*[local-name()='hlinkClick']")
	set linkLines to {}
	set seenURLs to {}
	repeat with hyperlinkNode in hyperlinkNodes
		set hyperlinkElement to contents of hyperlinkNode
		set propertyElement to hyperlinkElement's |parent|()
		set runElement to propertyElement's |parent|()
		set runName to (runElement's localName()) as text
		-- 只接受文字 run 或 field 直接持有的連結，排除段落預設樣式與形狀動作。
		if runName is "r" or runName is "fld" then
			set idNode to hyperlinkElement's attributeForLocalName:"id" URI:"http://schemas.openxmlformats.org/officeDocument/2006/relationships"
			if idNode is missing value then
				set relationID to ""
			else
				set relationID to (idNode's stringValue()) as text
			end if
			set targetURL to (my hyperlinkTarget(pageRels, relationID)) as text
			if targetURL is not "" and seenURLs does not contain targetURL then
				set textNodes to my nodes(runElement, ".//*[local-name()='t']")
				set visibleText to ""
				repeat with textNode in textNodes
					set visibleText to visibleText & (textNode's stringValue() as text)
				end repeat
				set visibleText to ((current application's NSString's stringWithString:visibleText)'s stringByTrimmingCharactersInSet:(current application's NSCharacterSet's whitespaceAndNewlineCharacterSet())) as text
				if visibleText is "" or visibleText is targetURL then
					set linkLine to ("[超連結] " & targetURL) as text
				else
					set linkLine to ("[超連結] " & visibleText & "：" & targetURL) as text
				end if
				set end of seenURLs to targetURL
				set end of linkLines to linkLine
			end if
		end if
	end repeat
	if (count of linkLines) = 0 then return ""
	set resultText to ""
	repeat with linkLine in linkLines
		set resultText to resultText & (contents of linkLine) & linefeed
	end repeat
	return resultText
end pageHyperlinks

on hyperlinkTarget(relsXML, relationID)
	set relNodes to my nodes(relsXML, "/*[local-name()='Relationships']/*[local-name()='Relationship']")
	repeat with relNode in relNodes
		if ((relNode's attributeForName:"Id")'s stringValue() as text) is relationID then
			set typeText to ((relNode's attributeForName:"Type")'s stringValue()) as text
			if typeText ends with "/hyperlink" then return ((relNode's attributeForName:"Target")'s stringValue()) as text
			return ""
		end if
	end repeat
	return ""
end hyperlinkTarget

-- 只取本頁 XML 真正引用的內嵌圖片，包含群組與形狀圖片填滿。
-- 保留原始圖片，OCR 可能讀到裁切範圍以外的字。
on pageImages(pagePath, packageRoot)
	set pageXML to my readXML(pagePath)
	set pageRels to my relationships(pagePath)
	set embeddedIDs to my nodes(pageXML, "//*[local-name()='blip']/@*[local-name()='embed']")
	set linkedIDs to my nodes(pageXML, "//*[local-name()='blip']/@*[local-name()='link']")
	set paths to {}
	repeat with idNode in embeddedIDs
		set mediaPath to my relatedPath(pageRels, (idNode's stringValue() as text), pagePath, packageRoot)
		if paths does not contain mediaPath then set end of paths to mediaPath
	end repeat
	if (linkedIDs's |count|() as integer) > 0 then
		set warningCount to warningCount + 1
		-- 不自動下載外部連結，也不把它誤報成辨識成功。
		set end of paths to "WARNING:此外本頁有外部連結圖片，未下載或辨識。"
	end if
	return paths
end pageImages

on recognizeImage(mediaPath, tempDir)
	if mediaPath starts with "WARNING:" then return "【警告】" & text 9 thru -1 of mediaPath
	set cachedText to ocrCache's objectForKey:mediaPath
	if cachedText is not missing value then return cachedText as text
	set ocrRunCount to ocrRunCount + 1
	set outputPath to tempDir & "/ocr-" & ocrRunCount & ".txt"
	set pasteboard to current application's NSPasteboard's generalPasteboard()
	set clipboardBackup to my copyClipboard(pasteboard)
	set markerText to "keynote-ocr-" & (current application's NSUUID's UUID()'s UUIDString() as text)
	try
		-- 現有捷徑的最後動作是通知，辨識結果在剪貼簿。
		-- 先放唯一標記，避免誤讀上一次 OCR 或使用者原本的文字。
		pasteboard's clearContents()
		pasteboard's setString:markerText forType:(current application's NSPasteboardTypeString)
		do shell script "/usr/bin/shortcuts run " & quoted form of shortcutName & " --input-path " & quoted form of mediaPath & " --output-path " & quoted form of outputPath & " --output-type public.utf8-plain-text"
		if (fm's fileExistsAtPath:outputPath) as boolean then
			set {ocrNSString, readError} to current application's NSString's stringWithContentsOfFile:outputPath encoding:(current application's NSUTF8StringEncoding) |error|:(reference)
			if ocrNSString is missing value then error "無法讀取捷徑的 UTF-8 文字輸出：" & (readError's localizedDescription() as text)
		else
			set ocrNSString to pasteboard's stringForType:(current application's NSPasteboardTypeString)
			if ocrNSString is missing value then error "捷徑未輸出文字檔，剪貼簿也沒有文字。"
			if (ocrNSString as text) is markerText then error "捷徑未輸出文字檔，也未更新剪貼簿。"
		end if
		set ocrText to (ocrNSString's stringByTrimmingCharactersInSet:(current application's NSCharacterSet's whitespaceAndNewlineCharacterSet())) as text
		my restoreClipboard(pasteboard, clipboardBackup)
		ocrCache's setObject:ocrText forKey:mediaPath
		return ocrText
	on error errorMessage number errorNumber
		my restoreClipboard(pasteboard, clipboardBackup)
		if errorNumber = -128 then error errorMessage number errorNumber
		set warningCount to warningCount + 1
		return "【OCR 失敗】" & errorMessage & " (" & errorNumber & ")"
	end try
end recognizeImage

-- 只比對完整的無文字回覆，避免誤刪含有其他辨識內容的結果。
on hasOCRText(ocrText)
	set trimCharacters to current application's NSCharacterSet's characterSetWithCharactersInString:(space & tab & return & linefeed & "（）()。.!！")
	set normalizedText to ((current application's NSString's stringWithString:ocrText)'s stringByTrimmingCharactersInSet:trimCharacters) as text
	return normalizedText is not in {"", "未辨識到文字", "圖片中沒有可辨識的文字", "此圖片中沒有文字", "圖片中沒有文字"}
end hasOCRText

-- 保存每個項目的所有可讀取資料型別，包含文字與圖片。
on copyClipboard(pasteboard)
	set savedItems to current application's NSMutableArray's array()
	set originalItems to pasteboard's pasteboardItems()
	if originalItems is missing value then return savedItems
	repeat with originalItem in originalItems
		set copiedItem to current application's NSPasteboardItem's alloc()'s init()
		repeat with typeName in originalItem's types()
			set itemData to originalItem's dataForType:typeName
			if itemData is not missing value then copiedItem's setData:itemData forType:typeName
		end repeat
		savedItems's addObject:copiedItem
	end repeat
	return savedItems
end copyClipboard

on restoreClipboard(pasteboard, savedItems)
	pasteboard's clearContents()
	if (savedItems's |count|() as integer) > 0 then pasteboard's writeObjects:savedItems
end restoreClipboard

on readXML(xmlPath)
	set xmlURL to current application's NSURL's fileURLWithPath:xmlPath
	set {xmlDocument, xmlError} to current application's NSXMLDocument's alloc()'s initWithContentsOfURL:xmlURL options:0 |error|:(reference)
	if xmlDocument is missing value then error "讀取 XML 失敗：" & xmlPath & "；" & (xmlError's localizedDescription() as text)
	return xmlDocument
end readXML

on nodes(xmlDocument, xpathText)
	set {matchedNodes, xpathError} to xmlDocument's nodesForXPath:xpathText |error|:(reference)
	if matchedNodes is missing value then error "XML 查詢失敗：" & (xpathError's localizedDescription() as text)
	return matchedNodes
end nodes

on relationships(partPath)
	set nsPath to current application's NSString's stringWithString:partPath
	set relsPath to (nsPath's stringByDeletingLastPathComponent() as text) & "/_rels/" & (nsPath's lastPathComponent() as text) & ".rels"
	return my readXML(relsPath)
end relationships

on relatedPath(relsXML, relationID, partPath, packageRoot)
	set relNodes to my nodes(relsXML, "/*[local-name()='Relationships']/*[local-name()='Relationship']")
	repeat with relNode in relNodes
		if ((relNode's attributeForName:"Id")'s stringValue() as text) is relationID then
			set modeNode to relNode's attributeForName:"TargetMode"
			if modeNode is not missing value then
				if (modeNode's stringValue() as text) is "External" then error "不支援外部圖片關聯：" & relationID
			end if
			set targetText to ((relNode's attributeForName:"Target")'s stringValue()'s stringByRemovingPercentEncoding()) as text
			if targetText starts with "/" then
				set resolvedPath to packageRoot & targetText
			else
				set resolvedPath to ((current application's NSString's stringWithString:partPath)'s stringByDeletingLastPathComponent() as text) & "/" & targetText
			end if
			set resolvedPath to ((current application's NSString's stringWithString:resolvedPath)'s stringByStandardizingPath()) as text
			if resolvedPath does not start with (packageRoot & "/") then error "關聯指向暫存簡報之外，已停止。"
			if not ((fm's fileExistsAtPath:resolvedPath) as boolean) then error "缺少關聯檔案：" & resolvedPath
			return resolvedPath
		end if
	end repeat
	error "找不到簡報關聯：" & relationID
end relatedPath

on writeUTF8(contentText, outputPath)
	set nsText to current application's NSString's stringWithString:contentText
	set {didWrite, writeError} to nsText's writeToFile:outputPath atomically:true encoding:(current application's NSUTF8StringEncoding) |error|:(reference)
	if not didWrite then error "寫入失敗：" & (writeError's localizedDescription() as text)
end writeUTF8

on absolutePath(pathText)
	set nsPath to (current application's NSString's stringWithString:pathText)'s stringByExpandingTildeInPath()
	if not (nsPath's isAbsolutePath() as boolean) then set nsPath to (fm's currentDirectoryPath())'s stringByAppendingPathComponent:nsPath
	return nsPath's stringByStandardizingPath() as text
end absolutePath

on unusedPath(proposedPath)
	set candidatePath to proposedPath
	set suffix to 2
	set basePath to (current application's NSString's stringWithString:proposedPath)'s stringByDeletingPathExtension() as text
	repeat while (fm's fileExistsAtPath:candidatePath) as boolean
		set candidatePath to basePath & "_" & suffix & ".txt"
		set suffix to suffix + 1
	end repeat
	return candidatePath
end unusedPath

on cleanup(tempDir)
	if tempDir is "" then return
	-- 僅移除 mktemp 建立的本次暫存資料夾。
	if tempDir starts with "/tmp/keynote-text." then fm's removeItemAtPath:tempDir |error|:(missing value)
end cleanup

on joinText(textList, separatorText)
	return (current application's NSArray's arrayWithArray:textList)'s componentsJoinedByString:separatorText as text
end joinText

on replaceText(sourceText, oldText, newText)
	return ((current application's NSString's stringWithString:sourceText)'s stringByReplacingOccurrencesOfString:oldText withString:newText) as text
end replaceText
