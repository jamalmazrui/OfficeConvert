; WdVert
; Version 1.0
; October 30, 2011
; Copyright 2011 by Jamal Mazrui
; GNU Lesser General Public License (LGPL)

#Include <File.au3>
#Include <Misc.au3>

Opt("MustDeclareVars", 1)
Dim $aSourceFiles, $aExtensions, $aFormats
Dim $bWordExisted, $bErrorEvent
Dim $i, $iFormat, $iSourceFile, $iTargetFormat, $iArgCount, $iConvertCount, $iSourceCount
Dim $s, $sProcess, $sSource, $sSourceDir, $sWildcards, $sSourceFile, $sSourceName, $sTargetFile, $sTargetFormat, $sTargetExtension, $sTargetDir, $sTarget
Dim $oErrorEvent, $oApp, $oDocs, $oDoc, $oExtensions

Const $vbTextCompare = 1
Const $msoEncodingUTF8 = 65001
Const $wdFormatDocument = 0
Const $wdFormatDocument97 = 0
Const $wdFormatTemplate = 1
Const $wdFormatTemplate97 = 1
Const $wdFormatText = 2
Const $wdFormatTextLineBreaks = 3
Const $wdFormatDOSText = 4
Const $wdFormatDOSTextLineBreaks = 5
Const $wdFormatRTF = 6
Const $wdFormatEncodedText = 7
Const $wdFormatUnicodeText = 7
Const $wdFormatHTML = 8
Const $wdFormatWebArchive = 9
Const $wdFormatFilteredHTML = 10
Const $wdFormatXML = 11
Const $wdFormatXMLDocument = 12
Const $wdFormatXMLDocumentMacroEnabled = 13
Const $wdFormatXMLTemplate = 14
Const $wdFormatXMLTemplateMacroEnabled = 15
Const $wdFormatOriginalFormatting = 16
Const $wdFormatDocumentDefault = 16
Const $wdFormatPDF = 17
Const $wdFormatXPS = 18
Const $wdFormatFlatXML = 19
Const $wdFormatFlatXMLMacroEnabled = 20
Const $wdFormatFlatXMLTemplate = 21
Const $wdFormatFlatXMLTemplateMacroEnabled = 22
Const $wdFormatPlainText = 22
Const $wdFormatOpenDocumentText = 23

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
$s = "Help for WdVert.exe -- Convert files using the API of Microsoft Word"
$s = $s & @CrLf & "Syntax:"
$s = $s & @CrLf & "WdVert Source Target TargetType"
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
$sProcess = "WinWord.exe"
$oExtensions = _CreateDictionary()
$oExtensions("doc") = $wdFormatDocument
$oExtensions("htm") = $wdFormatFilteredHTML
$oExtensions("html") = $wdFormatFilteredHTML
$oExtensions("pdf") = $wdFormatPDF
$oExtensions("rtf") = $wdFormatRTF
$oExtensions("txt") = $wdFormatText
$oExtensions("odt") = $wdFormatOpenDocumentText  
$oExtensions("xps") = $wdFormatXPS    
$oExtensions("mht") = $wdFormatWebArchive      
$oExtensions("mhtm") = $wdFormatWebArchive      
$oExtensions("xml") = $wdFormatXML        
$oExtensions("docx") = $wdFormatXMLDocument        

$iArgCount = $CmdLine[0]
If Not $iArgCount Then _HelpAndExit()

$sSource = $CmdLine[1]

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
If IsArray($aSourceFiles) Then $iSourceCount  = $aSourceFiles[0]

If ProcessExists($sProcess) Then $bWordexisted = True
$oApp = ObjCreate("Word.Application")
$oApp.Visible = False
$oApp.DisplayAlerts = False
$oApp.ScreenUpdating = False
$oDocs = $oApp.Documents

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

; Set $oDoc = $oDocs.Open($sSourceFile, AddToRecentFiles = False, ReadOnly = True, ConfirmConversions = False)
$oDoc = $oDocs.Open($sSourceFile, False, True, False)

If _FileExists($sTargetFile) Then _FileDelete($sTargetFile)
; $oDoc.SaveAs(FileName, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter, Encoding, InsertLineBreaks, AllowSubstitutions, LineEnding, AddBiDiMarks)
$oDoc.SaveAs($sTargetFile, $iTargetFormat, False, "", False, "", False, False, False, False, False, $msoEncodingUTF8  )
if $bErrorEvent then $bErrorEvent = 0
If _FileExists($sTargetFile) Then $iConvertCount = $iConvertCount + 1
$oDoc.Close()
$oDoc = 0
Next
$oDocs = 0

$oApp.Quit()
$oApp = 0
If Not $bWordExisted And ProcessExists($sProcess) Then ProcessClose($sProcess)


_ConsoleWriteLine("Converted " & $iConvertCount & " out of " & _StringPlural("file", $iSourceCount))
