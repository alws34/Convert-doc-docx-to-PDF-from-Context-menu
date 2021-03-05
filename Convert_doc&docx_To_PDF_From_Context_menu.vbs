option Explicit

'Determines if a string ends with the same characters as
'@param strValue the string to find the text
'@param checkFor the string to be searched for
'@param vbTextCompare for case-insenstive and vbBinaryCompare for case sensitive
'@return True if end with checkFor, false otherwise
'copied from http://www.freevbcode.com/ShowCode.asp?ID=2856
Public Function stringEndsWith(ByVal strValue,byVal checkFor, byVal compareType)
	Dim sCompare
	Dim lLen
	lLen = Len(checkFor)
	If lLen > Len(strValue) Then Exit Function
	sCompare = Right(strValue, lLen)
	stringEndsWith = StrComp(sCompare, checkFor, compareType) = 0
End Function

Function docToPdf(byVal docInputFile, byVal pdfOutputFile )
	Dim wordApplication
	Dim wordDocument
	Dim wordDocuments
	Const wdDoNotSaveChanges = 0
	Const wdFormatPDF        = 17
	'Dim currentDirectory

	'Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
	Set wordApplication = CreateObject("Word.Application")
	Set wordDocuments = wordApplication.Documents

	' Disable any potential macros of the word document.
	wordApplication.WordBasic.DisableAutoMacros
	WScript.StdOut.Write "opening "&docInputFile&"..."
	Set wordDocument = wordDocuments.Open(docInputFile)
	WScript.StdOut.Write "Done!"
	Wscript.StdOut.WriteLine
	
	WScript.StdOut.Write "Saving "&pdfOutputFile&"..."
	' See http://msdn2.microsoft.com/en-us/library/bb221597.aspx
	wordDocument.SaveAs pdfOutputFile, 17
	WScript.StdOut.Write "Done!"
	Wscript.StdOut.WriteLine
	
	wordDocument.Close WdDoNotSaveChanges
	wordApplication.Quit WdDoNotSaveChanges

	Set wordApplication = Nothing
End Function

'start of main program
Dim objFSO
Dim args
Dim arg
Dim filename
Dim outfilename
Dim strLen
Set objFSO = CreateObject("Scripting.FileSystemObject")

Set args = Wscript.Arguments

For Each arg In args
	Wscript.Echo arg
	if stringEndsWith(arg,"docx",vbTextCompare) then 
		filename = objFSO.GetAbsolutePathName(arg)
		strLen = len(filename)
		outfilename = left(filename,strLen-5) & ".pdf"
		docToPdf filename,outfilename
	elseif stringEndsWith(arg,"doc",vbTextCompare) then
		filename = objFSO.GetAbsolutePathName(arg)
		strLen = len(filename)
		outfilename = left(filename,strLen-4) & ".pdf"
		docToPdf filename,outfilename
	end if
Next