Set objFS = CreateObject("Scripting.FileSystemObject")

strFile = "C:\Users\ilh\Documents\BibTex\ThesisRef.bib"
strTemp = "C:\Users\ilh\Documents\BibTex\ThesisRefTemp.bib"

Set objFile = objFS.OpenTextFile(strFile)
Set objOutFile = objFS.CreateTextFile(strTemp,True)

strFind01 = "{\textless}sub{\textgreater}"
strReplace01 = "$_{"
strFind02 = "{\textless}/sub{\textgreater}"
strReplace02 = "}$"

Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    If InStr(strLine,strFind01)> 0 Then
        strLine = Replace(strLine,strFind01,strReplace01)
    End If 
	 If InStr(strLine,strFind02)> 0 Then
        strLine = Replace(strLine,strFind02,strReplace02)
    End If 
    ' WScript.Echo strLine
	 objOutFile.Write(strLine+vbCrLf)
Loop

objFile.Close
objOutFile.Close

objFS.DeleteFile(strFile)
objFS.MoveFile strTemp,strFile 