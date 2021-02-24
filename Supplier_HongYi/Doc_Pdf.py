from win32com.client import constants, gencache


def createPdf(wordPath):
    """
    word转pdf
    :param wordPath: word文件路径
    :param pdfPath:  生成pdf文件路径
    """
    word = gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(wordPath, ReadOnly=1)
    pdfPath = wordPath.rsplit(".", 1)[0] + ".pdf"
    doc.ExportAsFixedFormat(pdfPath,
                            constants.wdExportFormatPDF,
                            Item=constants.wdExportDocumentWithMarkup,
                            CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
    word.Quit(constants.wdDoNotSaveChanges)
    return pdfPath


# createPdf(r"C:\Users\windo\Desktop\【类型3】13家供应商\智微\13弘憶\IN2-20A0263 PK.doc")
