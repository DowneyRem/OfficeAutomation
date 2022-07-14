Attribute VB_Name = "Macros1"
Sub 打印()

    Application.PrintOut Range:=wdPrintAllDocument, Item:=wdPrintDocumentWithMarkup, Copies:=1, PageType:=wdPrintAllPages
    
End Sub
