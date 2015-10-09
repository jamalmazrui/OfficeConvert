; XlVert
; Version 1.0
; November 5, 2011
; Copyright 2011 by Jamal Mazrui
; GNU Lesser General Public License (LGPL)

#Include <File.au3>
#Include <Misc.au3>

Opt("MustDeclareVars", 1)
Dim $bExcelExisted, $bErrorEvent, $bNoteLabel, $bOutlineLabel, $bCommentLabel, $bHyperlinkLabel
Dim $iSlideCount, $iShapeCount, $iSlide, $iShape, $iCommentCount, $iComment, $iNoteCount, $iNote, $iHyperlinkCount, $iHyperlink
Dim $oErrorEvent, $oApp, $oXlss, $oXls, $oTextFrame, $oTextRange, $oShapes, $oShape, $oSlides, $oSlide, $oNotes, $oNote, $oComments, $oComment, $oHyperlinks, $oHyperlink
Dim $s, $sContents, $sBody, $sSourceFile, $sTargetFile, $sXls, $sText, $sName
Dim $aSourceFiles, $aExtensions, $aFormats
Dim $bExcelExisted, $bErrorEvent
Dim $i, $iFormat, $iSourceFile, $iTargetFormat, $iArgCount, $iConvertCount, $iSourceCount, $iSheetCount, $iSheet
Dim $s, $sProcess, $sSource, $sSourceDir, $sWildcards, $sSourceFile, $sSourceName, $sTargetFile, $sTargetFormat, $sTargetExtension, $sTargetDir, $sTarget
Dim $oErrorEvent, $oApp, $oXlss, $oXls, $oExtensions, $oLabels, $oSheets, $oSheet

Const $msoFalse = 0
Const $msoTrue = -1
Const $sDivider = @CrLf & "----------" & @CrLf & Chr(12) & @CrLf
Const $vbTextCompare = 1
Const $msoEncodingUTF8 = 65001

Const $msoAutomationSecurityLow = 1
Const $msoAutomationSecurityByUI = 2
Const $msoAutomationSecurityForceDisable = 3

Const $xlSYLK = 2
Const $xlWKS = 4
Const $xlWK1 = 5
Const $xlCSV = 6
Const $xlDBF2 = 7
Const $xlDBF3 = 8
Const $xlDIF = 9
Const $xlDBF4 = 11
Const $xlWJ2WD1 = 14
Const $xlWK3 = 15
Const $xlExcel2 = 16
Const $xlTemplate = 17
Const $xlAddIn = 18
Const $xlTextMac = 19
Const $xlTextWindows = 20
Const $xlTextMSDOS = 21
Const $xlCSVMac = 22
Const $xlCSVWindows = 23
Const $xlCSVMSDOS = 24
Const $xlIntlMacro = 25
Const $xlIntlAddIn = 26
Const $xlExcel2FarEast = 27
Const $xlWorks2FarEast = 28
Const $xlExcel3 = 29
Const $xlWK1FMT = 30
Const $xlWK1ALL = 31
Const $xlWK3FM3 = 32
Const $xlExcel4 = 33
Const $xlWQ1 = 34
Const $xlExcel4Workbook = 35
Const $xlTextPrinter = 36
Const $xlWK4 = 38
Const $xlExcel7 = 39
Const $xlWJ3 = 40
Const $xlWJ3FJ3 = 41
Const $xlUnicodeText = 42
Const $xlExcel9795 = 43
Const $xlHtml = 44
Const $xlWebArchive = 45
Const $xlXMLSpreadsheet = 46
Const $xlExcel12 = 50
Const $xlWorkbookDefault = 51
Const $xlOpenXMLWorkbook = 51
Const $xlOpenXMLWorkbookMacroEnabled = 52
Const $xlOpenXMLTemplateMacroEnabled = 53
Const $xlOpenXMLTemplate = 54
Const $xlOpenXMLAddIn = 55
Const $xlExcel8 = 56
Const $xlOpenDocumentSpreadsheet = 60
Const $xlWorkbookNormal = -4143
Const $xlCurrentPlatformText = -4158

Func _FileWriteUtf8b($sFile, $sBody)
Dim $iWriteMode = 2 + 128
FileDelete($sFile)
Dim $f = FileOpen($sFile, $iWriteMode)
Dim $iReturn = FileWrite($f, $sBody)
FileClose($f)
Return $iReturn
EndFunc

Func _ProcessComment($oComment, $sText)
Dim $sLabel = "Comment"
Dim $sReturn = $oComment.Text
If StringLower($sReturn) = StringLower($sLabel) Then $sReturn = ""
Dim $sAuthor = StringStripWS($oComment.Author, 3)
if $sReturn and $sAuthor Then $sReturn = "By " & $sAuthor & @CrLf & $sReturn

If Not $sReturn Then
$sReturn = $sText
ElseIf $oLabels($sLabel) Then
$sReturn = $sText & @CrLf & @CrLf & $sLabel & ":" & @CrLf & $sReturn
$oLabels($sLabel) = False
Else
$sReturn = $sText & @Crlf & @CrLf & $sReturn
EndIf
Return $sReturn
EndFunc

Func _ProcessHyperlink($oHyperlink, $sText)
Dim $sLabel = "Hyperlink"
Dim $sAddress = StringStripWS($oHyperlink.Address, 3)
Dim $sScreenTip = StringStripWS($oHyperlink.ScreenTip, 3)
Dim $sTextToDisplay = StringStripWS($oHyperlink.TextToDisplay, 3)
Dim $sReturn = $sAddress
If $sReturn And $sTextToDisplay Then $sReturn = $sTextToDisplay & @CrLf & $sReturn
If $sReturn And $sScreenTip Then $sReturn = $sScreenTip & @CrLf & $sReturn
If StringLower($sReturn) = StringLower($sLabel) Then $sReturn = ""

If Not $sReturn Then
$sReturn = $sText
ElseIf $oLabels($sLabel) Then
$sReturn = $sText & @CrLf & @CrLf & $sLabel & ":" & @CrLf & $sReturn
$oLabels($sLabel) = False
Else
$sReturn = $sText & @Crlf & @CrLf & $sReturn
EndIf
Return $sReturn
EndFunc

Func _ProcessShape($oShape, $sText, $sLabel)
Dim $oTextFrame, $oTextRange, $oTextEffect
Dim $sReturn = ""
Dim $sTextRange = ""
Dim $sTextEffect = ""
Dim $sAlternativeText = ""

If $oShape.HasTextFrame Then
$oTextFrame = $oShape.TextFrame
$oTextRange = $oTextFrame.TextRange
$sTextRange = StringStripWS($oTextRange.Text, 3)
EndIf

Dim $iType = $oShape.type
If $iType = 15 Then ; TextEffect
$oTextEffect = $oShape.TextEffect
$sTextEffect = StringStripWS($oTextEffect.Text, 3)
EndIf

$sAlternativeText = StringStripWS($oShape.AlternativeText, 3)

If StringLower($sTextRange) <> StringLower($sLabel) Then $sReturn = $sTextRange
If $sTextEffect And Not StringInStr($sReturn, $sTextEffect) Then $sReturn = $sReturn & @CrLf & $sTextEffect
If $sAlternativeText And Not StringInStr($sReturn, $sAlternativeText) Then $sReturn = $sReturn & @CrLf & $sAlternativeText
$sReturn = StringStripWS($sReturn, 3)
If StringIsInt($sReturn) Then $sReturn = ""

If Not $sReturn Then
$sReturn = $sText
ElseIf $oLabels($sLabel) Then
$sReturn = $sText & @CrLf & @CrLf & $sLabel & ":" & @CrLf & $sReturn
$oLabels($sLabel) = False
Else
$sReturn = $sText & @Crlf & @CrLf & $sReturn
EndIf

$oTextEffect = 0
$oTextRange = 0
$oTextFrame = 0

Return $sReturn
EndFunc

Func _ConsoleWriteLine($sLine)
Return ConsoleWrite($sLine & @CrLf)
EndFunc

Func _ErrorFunc()
Dim $sText
$sText = "Error" & @CrLf
$sText = $sText & "Number: " & Hex($oErrorEvent.Number, 8) & @CrLf
$sText = $sText & "WinDescription: " & $oErrorEvent.WinDescription & @CrLf
$sText = $sText & "Source: " & $oErrorEvent.Source & @CrLf
$sText = $sText & "Description: " & $oErrorEvent.Description & @CrLf
$sText = $sText & "HelpFile: " & $oErrorEvent.HelpFile & @CrLf
$sText = $sText & "HelpContext: " & $oErrorEvent.HelpContext & @CrLf
$sText = $sText & "LastDLLError: " & $oErrorEvent.LastDLLError & @CrLf
$sText = $sText & "ScriptLine: " & $oErrorEvent.ScriptLine & @CrLf
_ConsoleWriteLine($sText)
; _Show($sText)
$oErrorEvent.Clear()
$bErrorEvent = 1
EndFunc

Func _HelpAndExit()
Dim $s
$s = "Help for XlVert.exe -- Convert files using the API of Microsoft Excel"
$s = $s & @CrLf & "Syntax:"
$s = $s & @CrLf & "XlVert Source Target TargetType"
$s = $s & @CrLf & "where Source is the path to a file, directory, or wildcard specification"
$s = $s & @CrLf & "optional Target is the path to either a file or directory, defaulting to the source directory"
$s = $s & @CrLf & "optional TargetType is the target file type, as indicated by a common extension, integer constant, or string constant, defaulting to the txt extension"
_ConsoleWriteLine($s)
Exit
EndFunc

Func _Show($sTitle = "", $sMessage = "")
Return MsgBox(4096, $sTitle, $sMessage)
EndFunc

Func _CreateDictionary()
Dim $oDictionary
$oDictionary = ObjCreate("Scripting.Dictionary")
$oDictionary.CompareMode = $vbTextCompare
Return $oDictionary
EndFunc

Func _FileDelete($sFile)
; Delete a file if it exists, and test whether it is subsequently absent
; either because it was successfully deleted or because it was not present in the first place

Dim $oSystem

If Not _FileExists($sFile) Then Return True

$oSystem =ObjCreate("Scripting.FileSystemObject")
$oSystem.DeleteFile($sFile, True)
Return Not _FileExists($sFile)
EndFunc

Func _FileExists($sFile)
; Test whether File exists

Dim $oSystem

$oSystem =ObjCreate("Scripting.FileSystemObject")
Return Not $oSystem.FolderExists($sFile) And $oSystem.FileExists($sFile)
EndFunc

Func _DirExists($sDir)
; Test whether directory exists

Dim $oSystem

$oSystem =ObjCreate("Scripting.FileSystemObject")
Return $oSystem.FolderExists($sDir)
EndFunc

Func _PathCombine($sDir, $sName)
; Combine directory and name to form a valid path

Dim $sPath

$sPath = $sDir & "\" & $sName
Return StringReplace($sPath, "\\", "\")
EndFunc

Func _PathGetBase($sPath)
; Get base/root name of a file or directory

Dim $oSystem

$oSystem =ObjCreate("Scripting.FileSystemObject")
Return $oSystem.GetBaseName($sPath)
EndFunc

Func _PathGetExtension($sPath)
; Get extention of file or directory

Dim $oSystem

$oSystem =ObjCreate("Scripting.FileSystemObject")
Return $oSystem.GetExtensionName($sPath)
EndFunc

Func _PathGetDir($sPath)
; Get the parent directory of a file or directory

Dim $oSystem
$oSystem =ObjCreate("Scripting.FileSystemObject")
Return $oSystem.GetParentFolderName($sPath)
EndFunc

Func _PathGetName($sPath)
; Get the file or directory name at the end of a path

Dim $oSystem

$oSystem =ObjCreate("Scripting.FileSystemObject")
Return $oSystem.GetFileName($sPath)
EndFunc

Func _StringPlural($sItem, $iCount)
; Return singular or plural form of a string, depending on whether count equals one

Dim $sReturn

$sReturn = $iCount & " " & $sItem
If $iCount <> 1 Then $sReturn = $sReturn & "s"
Return $sReturn
EndFunc

; Main program
$oErrorEvent = ObjEvent("AutoIt.Error","_ErrorFunc")
$sProcess = "Excel.exe"
$oLabels = _CreateDictionary()

$oExtensions = _CreateDictionary()
$oExtensions("xls") = $xlWorkbookDefault
$oExtensions("csv") = $xlCSV
; $oExtensions("dbf") = $xlDBF3
$oExtensions("dif") = $xlDIF
$oExtensions("htm") = $xlHtml
$oExtensions("html") = $xlHtml
$oExtensions("sylk") = $xlSYLK
; $oExtensions("pdf") = $ppSaveAsPDF
; $oExtensions("rtf") = $ppSaveAsRTF
$oExtensions("txt") = $xlTextWindows
$oExtensions("ods") = $xlOpenDocumentSpreadsheet
; $oExtensions("xps") = $ppSaveAsXPS
$oExtensions("mht") = $xlWebArchive
$oExtensions("mhtm") = $xlWebArchive
$oExtensions("xml") = $xlXMLSpreadsheet
; $oExtensions("xml") = $xlOpenXMLWorkbook
$oExtensions("wks") = $xlWKS

$iArgCount = $CmdLine[0]
If Not $iArgCount Then _HelpAndExit()

$sSource = $CmdLine[1]
; $sSource = "*.xls"

$sSourceDir = $sSource
If Not _DirExists($sSourceDir) Then $sSourceDir = _PathGetDir($sSource)
$sSource = _PathFull($sSource)
If Not $sSourceDir Then $sSourceDir = @WorkingDir

$sTarget = $sSourceDir
If $iArgCount > 1 Then $sTarget = $CmdLine[2]
$sTarget = _PathFull($sTarget)
$sTargetDir = $sTarget
If Not _DirExists($sTargetDir) Then $sTargetDir = _PathGetDir($sTargetDir)

$sTargetExtension = "txt"
If Not _DirExists($sTarget) Then $sTargetExtension = _PathGetExtension($sTarget)
If $iArgCount > 2 Then $sTargetExtension = $CmdLine[3]
$iTargetFormat = -1
If $oExtensions.Exists($sTargetExtension) Then
$iTargetFormat = $oExtensions($sTargetExtension)
ElseIf StringIsDigit($sTargetExtension) Then
$iTargetFormat = Number($sTargetExtension)
$sTargetExtension = ""
Else
$iTargetFormat = Eval($sTargetExtension)
If IsString($iTargetFormat) And Not StringLen($iTargetFormat) Then $iTargetFormat = Eval("wdFormat" & $sTargetExtension)
$sTargetExtension = ""
EndIf

If Not $sTargetExtension Then
$aExtensions = $oExtensions.Keys
$aFormats = $oExtensions.Items
For $i = 0 To UBound($aFormats) - 1
$iFormat = $aFormats[$i]
If $iFormat = $iTargetFormat Then
$sTargetExtension = $aExtensions[$i]
ExitLoop
EndIf
Next
EndIf

$sWildCards = "*.*"
If Not _DirExists($sSource) Then $sWildcards = _PathGetName($sSource)
$aSourceFiles = _FileListToArray($sSourceDir, $sWildcards, 1)
$iSourceCount = 0
If IsArray($aSourceFiles) Then $iSourceCount = $aSourceFiles[0]

If ProcessExists($sProcess) Then $bExcelexisted = True
$oApp = ObjCreate("Excel.Application")
$oApp.AutomationSecurity = $msoAutomationSecurityLow
$oApp.DisplayAlerts = False
$oApp.ScreenUpdating = False
; $oApp.Visible = True ; Needed for automation to work
$oApp.Visible = False

$oXlss = $oApp.WorkBooks
; _Show("oXlss", IsObj($oXlss))

$iConvertCount = 0
For $iSourceFile = 1 To $iSourceCount
If $iSourceFile = 1 Then _ConsoleWriteLine("Converting")
$sSourceFile = $aSourceFiles[$iSourceFile]
$sSourceFile = _PathCombine($sSourceDir, $sSourceFile)
$s = $sTargetExtension
If $s Then $s = "." & $s
$sTargetFile = $sTarget
If _DirExists($sTargetFile) Then $sTargetFile = _PathCombine($sTargetDir, _PathGetBase($sSourceFile) & $s)
If StringLower($sSourceFile) = StringLower($sTargetFile) Then ContinueLoop

$sSourceName = _PathGetName($sSourceFile)
_ConsoleWriteLine($sSourceName)

; expression.Open(FileName, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru, Local, CorruptLoad)
$oXls = $oXlss.Open($sSourceFile)
; _Show("oXls", IsObj($oXls))

If _FileExists($sTargetFile) Then _FileDelete($sTargetFile)
; _Show($sTargetExtension)
If $sTargetExtension = "txt" Then
$oSheets = $oXls.sheets
$iSheetCount =$oSheets.Count
$sText = ""
$sContents = "Contents" & @CrLf
For $iSheet = 1 To $iSheetCount
$oSheet = $oSheets.Item($iSheet)
If _FileExists($sTargetFile) Then _FileDelete($sTargetFile)
$oSheet.SaveAs($sTargetFile, $xlTextWindows)
$sName = StringStripWS($oSheet.Name, 3)
If $iSheetCount > 1 Then $sName = "Sheet " & String($iSheet) & ": " & $sName
$sBody = FileRead($sTargetFile)
$sBody = $sName & @CrLf & @CrLf & $sBody
If $sText Then $sText = $sText & $sDivider
$sText = $sText & $sBody
$sContents = $sContents & @CrLf & $sName
$oSheet = 0
$oXls.Close()
$oXls = 0
$oXls = $oXlss.Open($sSourceFile)
$oSheets = $oXls.Sheets
Next

If $iSheetCount > 1 Then $sText = String($iSheetCount) & "Sheets" & $sContents & @CrLf & $sText
; $sText = StringRegExpReplace($sText, "\r\n\s*?\r\n\s*?(\r\n\s*?)+", "\r\n\r\n")
; $sText = StringRegExpReplace($sText, "\s\s\s\s)\s+", "$1")
$sText = StringStripWS($sText, 3) & @CrLf
; _Show("Error " & String(@Error), "Extended " & String(@Extended))
_FileWriteUtf8b($sTargetFile, $sText)
Else
; _Show("IsObject", IsObj($oXls))
; expression.SaveAs(FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout, Local)
$oXls.SaveAs($sTargetFile, $iTargetFormat)
; _Show($bErrorEvent, $sTargetFile)
if $bErrorEvent then $bErrorEvent = 0
EndIf

If _FileExists($sTargetFile) Then $iConvertCount = $iConvertCount + 1
$oXls.Close()
$oXls = 0
Next
$oXlss = 0

$oApp.Quit()
$oApp = 0
If Not $bExcelExisted And ProcessExists($sProcess) Then ProcessClose($sProcess)

_ConsoleWriteLine("Converted " & $iConvertCount & " out of " & _StringPlural("file", $iSourceCount))
