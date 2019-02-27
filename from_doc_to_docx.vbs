convert_doc_to_docx

Sub convert_doc_to_docx()
    Set oWord = CreateObject("Word.Application")
    Set word_document = oWord.Documents.Open( "D:\projects\acts_preprocessing_with_nlp" & _
            "\data\corpus\14-AF-3301-03251-CV.doc")
    word_document.SaveAs2 "D:\projects\acts_preprocessing_with_nlp" & _
            "\data\corpus\14-AF-3301-03251-CV.docx", 12
    word_document.close()
End Sub