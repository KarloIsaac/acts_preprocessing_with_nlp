'Opens a MS Word file and saves it with the .docx extension.
'
'Parameters
'--------
'file_path
'   The full path to the file to the reference file. It must be provided as one of the scripts
'   parameters
'
' Notes
' -------
' The new file with the .docx extension is saved in the same directory as the original file.
' This script doesn't verify what is the extension of the file it is opening. Likewise, it does 
' not verify whether a file with the same name alrready exists in the directory. It is 
' responsability of the caller to account for this possibilities and adjust its behaviour 
' accordingly

convert_doc_to_docx

Sub convert_doc_to_docx()
    file_path = Wscript.Arguments(0)
    Set oWord = CreateObject("Word.Application")
    Set word_document = oWord.Documents.Open(file_path)
    word_document.SaveAs2 file_path & "x", 12
    word_document.close()
    oWord.Quit
End Sub