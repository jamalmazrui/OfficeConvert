; PpVert
; Version 1.0
; November 5, 2011
; Copyright 2011 by Jamal Mazrui
; GNU Lesser General Public License (LGPL)

#Include <File.au3>
#Include <Misc.au3>

Opt("MustDeclareVars", 1)
Dim $bPowerPointExisted, $bErrorEvent, $bNoteLabel, $bOutlineLabel, $bCommentLabel, $bHyperlinkLabel
Dim $iSlideCount, $iShapeCount, $iSlide, $iShape, $iCommentCount, $iComment, $iNoteCount, $iNote, $iHyperlinkCount, $iHyperlink
Dim $oErrorEvent, $oApp, $oPpts, $oPpt, $oTextFrame, $oTextRange, $oShapes, $oShape, $oSlides, $oSlide, $oNotes, $oNote, $oComments, $oComment, $oHyperlinks, $oHyperlink
Dim $s, $sSourceFile, $sTargetFile, $sPpt, $sText, $oLabels

Dim $aSourceFiles, $aExtensions, $aFormats
Dim $bPowerPointExisted, $bErrorEvent
Dim $i, $iFormat, $iSourceFile, $iTargetFormat, $iArgCount, $iConvertCount, $iSourceCount
Dim $s, $sProcess, $sSource, $sSourceDir, $sWildcards, $sSourceFile, $sSourceName, $sTargetFile, $sTargetFormat, $sTargetExtension, $sTargetDir, $sTarget
Dim $oErrorEvent, $oApp, $oPpts, $oPpt, $oExtensions

Const $msoFalse = 0
Const $msoTrue = -1
Const $sDivider = @CrLf & "----------" & @CrLf & Chr(12) & @CrLf
Const $vbTextCompare = 1
Const $msoEncodingUTF8 = 65001

Const $msoAutomationSecurityLow = 1
Const $msoAutomationSecurityByUI = 2
Const $msoAutomationSecurityForceDisable = 3

Const $ppSaveAsPresentation = 1
Const $ppSaveAsText = 2
Const $ppSaveAsTemplate = 5
Const $ppSaveAsRTF = 6
Const $ppSaveAsShow = 7
Const $ppSaveAsAddIn = 8
Const $ppSaveAsDefault = 11
Const $ppSaveAsHTML = 12
Const $ppSaveAsHTMLv3 = 13
Const $ppSaveAsHTMLDual = 14
Const $ppSaveAsMetaFile = 15
Const $ppSaveAsGIF = 16
Const $ppSaveAsJPG = 17
Const $ppSaveAsPNG = 18
Const $ppSaveAsBMP = 19
Const $ppSaveAsWebArchive = 20
Const $ppSaveAsTIF = 21
Const $ppSaveAsEMF = 23
Const $ppSaveAsOpenXMLPresentation = 24
Const $ppSaveAsOpenXMLPresentationMacroEnabled = 25
Const $ppSaveAsOpenXMLTemplate = 26
Const $ppSaveAsOpenXMLTemplateMacroEnabled = 27
Const $ppSaveAsOpenXMLShow = 28
Const $ppSaveAsOpenXMLShowMacroEnabled = 29
Const $ppSaveAsOpenXMLAddin = 30
Const $ppSaveAsOpenXMLTheme = 31
Const $ppSaveAsPDF = 32
Const $ppSaveAsXPS = 33
Const $ppSaveAsXMLPresentation = 34
Const $ppSaveAsOpenDocumentPresentation = 35
Const $ppSaveAsExternalConverter = 36

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
$s = "Help for PpVert.exe -- Convert files using the API of Microsoft PowerPoint"
$s = $s & @CrLf & "Syntax:"
$s = $s & @CrLf & "PpVert Source Target TargetType"
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
$sProcess = "PowerPnt.exe"
$oLabels = _CreateDictionary()

$oExtensions = _CreateDictionary()
$oExtensions("ppt") = $ppSaveAsPresentation
$oExtensions("htm") = $ppSaveAsHTML
$oExtensions("html") = $ppSaveAsHTML
$oExtensions("pdf") = $ppSaveAsPDF
$oExtensions("rtf") = $ppSaveAsRTF
$oExtensions("txt") = $ppSaveAsText
$oExtensions("odp") = $ppSaveAsOpenDocumentPresentation
$oExtensions("xps") = $ppSaveAsXPS
$oExtensions("mht") = $ppSaveAsWebArchive
$oExtensions("mhtm") = $ppSaveAsWebArchive
$oExtensions("xml") = $ppSaveAsOpenXMLPresentation
$oExtensions("pptx") = $ppSaveAsXMLPresentation

$iArgCount = $CmdLine[0]
If Not $iArgCount Then _HelpAndExit()

$sSource = $CmdLine[1]
; $sSource = "c:\broadband\*.pp*"

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
If IsString($iTargetFormat) And Not StringLen($iTargetFormat) Then $iTargetFormat = Eval("ppSaveAs" & $sTargetExtension)
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

If ProcessExists($sProcess) Then $bPowerPointexisted = True
$oApp = ObjCreate("PowerPoint.Application")
$oApp.AutomationSecurity = $msoAutomationSecurityLow
$oApp.DisplayAlerts = False
$oApp.Visible = True ; Needed for automation to work 

$oPpts = $oApp.Presentations
; _Show("oPpts", IsObj($oPpts))

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

; expression.Open(FileName, ReadOnly, Untitled, WithWindow)
$oPpt = $oPpts.Open($sSourceFile, $msoTrue, $msoTrue, $msoFalse)
; _Show("oPpt", IsObj($oPpt))

If _FileExists($sTargetFile) Then _FileDelete($sTargetFile)
; _Show($sTargetExtension)
If $sTargetExtension = "txt" Then
$sPpt = $oPpt.name
$sText = _PathGetBase($sPpt)

$oSlides = $oPpt.slides
$iSlideCount = $oSlides.Count
$sText = $sText & @CrLf & _StringPlural("Slide", $iSlideCount)
For $iSlide = 1 To $iSlideCount
$oLabels("Comments") = True
$oLabels("Hyperlinks") = True
$oLabels("Notes") = True
; $oLabels("Outline") = True
$oLabels("Outline") = False ; Just assume that text after slide heading is outline content

$oSlide = $oSlides.Item($iSlide)
$sText = $sText & $sDivider & "Slide " & String($iSlide)

$oShapes = $oSlide.Shapes
$iShapeCount = $oShapes.Count
for $iShape = 1 to $iShapeCount
$oShape = $oShapes.Item($iShape)
$sText = _ProcessShape($oShape, $sText, "Outline")
$oShape = 0
Next
$oShapes = 0

$oNotes = $oSlide.NotesPage
$iNoteCount = $oNotes.Count
; If $iNoteCount > 1 Then _Show($iNoteCount)
for $iNote = 1 to $iNoteCount
$oNote = $oNotes.Item($iNote)
$oShapes = $oNote.Shapes
$iShapeCount = $oShapes.Count
For $iShape = 1 To $iShapeCount
$oShape = $oShapes.Item($iShape)
$sText = _ProcessShape($oShape, $sText, "Notes")
$oShape = 0
Next
$oShapes = 0
$oNote = 0
Next
$oNotes = 0

$oComments = $oSlide.Comments
$iCommentCount = $oComments.Count
$bCommentLabel = True
for $iComment = 1 to $iCommentCount
$oComment = $oComments.Item($iComment)
$sText = _ProcessComment($oComment, $sText)
$oComment = 0
Next
$oComments = 0
$oHyperlinks = $oSlide.Hyperlinks
$iHyperlinkCount = $oHyperlinks.Count
$bHyperlinkLabel = True
for $iHyperlink = 1 to $iHyperlinkCount
$oHyperlink = $oHyperlinks.Item($iHyperlink)
$sText = _ProcessHyperlink($oHyperlink, $sText)
$oHyperlink = 0
Next
$oHyperlinks = 0

$oSlide = 0
Next
$oSlides = 0
_FileWriteUtf8b($sTargetFile, $sText)
Else
; expression.SaveAs(Filename, FileFormat, EmbedFonts)
; _Show("IsObject", IsObj($oPpt))
$oPpt.SaveAs($sTargetFile, $iTargetFormat)
; _Show($bErrorEvent, $sTargetFile)
if $bErrorEvent then $bErrorEvent = 0
EndIf

If _FileExists($sTargetFile) Then $iConvertCount = $iConvertCount + 1
$oPpt.Close()
$oPpt = 0
Next
$oPpts = 0

$oApp.Quit()
$oApp = 0
If Not $bPowerPointExisted And ProcessExists($sProcess) Then ProcessClose($sProcess)


_ConsoleWriteLine("Converted " & $iConvertCount & " out of " & _StringPlural("file", $iSourceCount))
