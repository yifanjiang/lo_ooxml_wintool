'************************************************
'
' DOC2PDF.VBS Microsoft Scripting Host Script (Requires Version 5.6 or newer)
' --------------------------------------------------------------------------------
'
' Author: Michael Suodenjoki
' Created: 2007.07.07
'
' This script can create a PDF file from a Word document provided you're using
' Word 2007 and have the 'Office Add-in: Save As PDF' installed.
'

' Modified by Yifan Jiang 2013.02.27
' To print out Powerpoint and Excel documents as well.
' Usage:
'   cscript /nologo "ms2pdf.vbs" /nologo $msoffice_file_name [/o:<pdf-file>]

' Constants
Const WdDoNotSaveChanges = 0
' see WdSaveFormat enumeration constants: 
' http://msdn2.microsoft.com/en-us/library/bb238158.aspx
Const wdFormatPDF = 17  ' Word PDF format.
Const wdFormatXPS = 18  ' Word XPS format
Const ppSaveAsPDF = 32  ' Powerpoint PDF format.
Const ppSaveAsXPS = 33  ' Powerpoint XPS format
Const xlTypePDF = 0     ' Excel PDF format
Const xlTypeXPS = 1     ' Excel SPS format

' Global variables
Dim arguments
Set arguments = WScript.Arguments

' ***********************************************
' ECHOLOGO
'
' Outputs the logo information.
'
Function EchoLogo()
  If Not (arguments.Named.Exists("nologo") Or arguments.Named.Exists("n")) Then
    WScript.Echo "doc2pdf Version 2.0, Michael Suodenjoki 2007"
    WScript.Echo "=================================================="
    WScript.Echo ""
  End If
End Function

' ***********************************************
' ECHOUSAGE
'
' Outputs the usage information.
'
Function EchoUsage()
  If arguments.Count=0 Or arguments.Named.Exists("help") Or _
    arguments.Named.Exists("h") _
  Then
    WScript.Echo "Generates a PDF from a Word document file using Word 2007."
    WScript.Echo ""
    WScript.Echo "Usage: doc2pdf.vbs <options> <doc-file> [/o:<pdf-file>]"
    WScript.Echo ""
    WScript.Echo "Available Options:"
    WScript.Echo ""
    WScript.Echo " /nologo - Specifies that the logo shouldn't be displayed"
    WScript.Echo " /help   - Specifies that this usage/help information " + _
                 "should be displayed."
    WScript.Echo " /debug  - Specifies that debug output should be displayed."
    WScript.Echo ""
    WScript.Echo "Parameters:"
    WScript.Echo ""
    WScript.Echo " /o:<pdf-file> Optionally specification of output file (PDF)."
    WScript.Echo ""
  End If 
End Function

' ***********************************************
' CHECKARGS
'
' Makes some preliminary checks of the arguments.
' Quits the application is any problem is found.
'
Function CheckArgs()
  ' Check that <doc-file> is specified
  If arguments.Unnamed.Count <> 1 Then
    WScript.Echo "Error: Obligatory <doc-file> parameter missing!"
    WScript.Quit 1
  End If

  bShowDebug = arguments.Named.Exists("debug") Or arguments.Named.Exists("d")

End Function


' ***********************************************
' DOC2PDF
'
' Converts a Word document to PDF using Word 2007.
'
' Input:
' sDocFile - Full path to Word document.
' sPDFFile - Optional full path to output file.
'
' If not specified the output PDF file
' will be the same as the sDocFile except
' file extension will be .pdf.
'

Dim fso ' As FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

Function DOC2PDF( sDocFile, sFolder )

  Dim wdo ' As Word.Application
  Dim wdoc ' As Word.Document
  Dim wdocs ' As Word.Documents
  Dim sPrevPrinter ' As String

  Set wdo = CreateObject("Word.Application")
  Set wdocs = wdo.Documents

  sDocFile = fso.GetAbsolutePathName(sDocFile)
  sFolder = fso.GetAbsolutePathName(sFolder)

  If Len(sFolder)=0 Then

    sFolder = fso.GetParentFolderName(sDocFile)

  End If

  sPDFFile = sFolder + "\" + fso.GetFileName(sDocFile) + ".pdf"

  ' Enable this line if you want to disable autoexecute macros
  ' wdo.WordBasic.DisableAutoMacros

  ' Open the Word document
  Set wdoc = wdocs.Open(sDocFile)

  ' ' Debug outputs...

    ' WScript.Echo "Doc file = '" + sDocFile + "'"
    ' WScript.Echo "PDF file = '" + sPDFFile + "'"



  ' Let Word document save as PDF
  ' - for documentation of SaveAs() method,
  '   see http://msdn2.microsoft.com/en-us/library/bb221597.aspx 
  wdoc.SaveAs sPDFFile, wdFormatPDF

  wdoc.Close WdDoNotSaveChanges
  wdo.Quit WdDoNotSaveChanges
  Set wdo = Nothing

  Set fso = Nothing

End Function

Function PPT2PDF( sDocFile, sFolder )

  Dim wdo ' As Word.Application
  Dim wdoc ' As Word.Document
  Dim wdocs ' As Word.Documents
  Dim sPrevPrinter ' As String

  Set wdo = CreateObject("Powerpoint.Application")
  Set wdocs = wdo.Presentations

  sDocFile = fso.GetAbsolutePathName(sDocFile)
  sFolder = fso.GetAbsolutePathName(sFolder)

  If Len(sFolder)=0 Then

    sFolder = fso.GetParentFolderName(sDocFile)

  End If

  sPDFFile = sFolder + "\" + fso.GetFileName(sDocFile) + ".pdf"

  ' Enable this line if you want to disable autoexecute macros
  ' wdo.WordBasic.DisableAutoMacros

  ' Open the Word document
  Set wdoc = wdocs.Open(sDocFile)

  ' Let Word document save as PDF
  ' - for documentation of SaveAs() method,
  '   see http://msdn2.microsoft.com/en-us/library/bb221597.aspx 
  wdoc.SaveAs sPDFFile, ppSaveAsPDF

  wdoc.Close
  wdo.Quit

  Set wdo = Nothing
  Set fso = Nothing

End Function

Function XLS2PDF( sDocFile, sFolder )

  Dim wdo ' As Word.Application
  Dim wdoc ' As Word.Document
  Dim wdocs ' As Word.Documents
  Dim sPrevPrinter ' As String

  Set wdo = CreateObject("Excel.Application")
  Set wdocs = wdo.Workbooks

  sDocFile = fso.GetAbsolutePathName(sDocFile)
  sFolder = fso.GetAbsolutePathName(sFolder)

  If Len(sFolder)=0 Then

    sFolder = fso.GetParentFolderName(sDocFile)

  End If

  sPDFFile = sFolder + "\" + fso.GetFileName(sDocFile) + ".pdf"

  ' ' Debug outputs...
  ' If bShowDebug Then
  '   WScript.Echo "Doc file = '" + sDocFile + "'"
  '   WScript.Echo "PDF file = '" + sPDFFile + "'"
  ' End If

  ' If Len(sPDFFile)=0 Then
  '   sPDFFile = fso.GetFileName(sDocFile) + ".pdf"
  ' End If

  ' If Len(fso.GetParentFolderName(sPDFFile))=0 Then
  '   sPDFFile = sFolder + "\" + sPDFFile
  ' End If

  ' Enable this line if you want to disable autoexecute macros
  ' wdo.WordBasic.DisableAutoMacros

  ' Open the Word document
  Set wdoc = wdocs.Open(sDocFile)

  ' Let Word document save as PDF
  ' - for documentation of SaveAs() method,
  '   see http://msdn2.microsoft.com/en-us/library/bb221597.aspx 
  ' wdoc.SaveAs sPDFFile, wdFormatPDF

  wdoc.ExportAsFixedFormat xlTypePDF, sPDFFile

' wdoc.ExportAsFixedFormat Type:=xlTypePDF, Filename:=sPDFFile, Quality:=xlQualityStandard

  wdoc.Close
  wdo.Quit

  Set wdo = Nothing
  Set fso = Nothing

End Function

' *** MAIN **************************************

Call EchoLogo()
Call EchoUsage()
Call CheckArgs()

Dim sFileExt : sFileExt = UCase(fso.GetExtensionName(arguments.Unnamed.Item(0)))

Select Case sFileExt

    Case "DOC"

        WScript.Echo arguments.Unnamed.Item(0)

        Call DOC2PDF( arguments.Unnamed.Item(0), arguments.Named.Item("o") )

    Case "DOCX"

        WScript.Echo arguments.Unnamed.Item(0)

        Call DOC2PDF( arguments.Unnamed.Item(0), arguments.Named.Item("o") )

    Case "XLS"

        WScript.Echo arguments.Unnamed.Item(0)

        Call XLS2PDF( arguments.Unnamed.Item(0), arguments.Named.Item("o") )

    Case "XLSX"

        WScript.Echo arguments.Unnamed.Item(0)

        Call XLS2PDF( arguments.Unnamed.Item(0), arguments.Named.Item("o") )

    Case "PPT"

        WScript.Echo arguments.Unnamed.Item(0)

        Call PPT2PDF( arguments.Unnamed.Item(0), arguments.Named.Item("o") )

    Case "PPTX"

        WScript.Echo arguments.Unnamed.Item(0)

        Call PPT2PDF( arguments.Unnamed.Item(0), arguments.Named.Item("o") )

    Case Else

        WScript.Echo "Format not supported."

End Select

Set arguments = Nothing