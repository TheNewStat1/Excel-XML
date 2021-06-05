using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OpenCML_Sandbox
{
    class XLWorkBook : SpreadsheetDocument
    {
        new public static SpreadsheetDocument Create(string filePath, SpreadsheetDocumentType documentType)
        {
            SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
            return document;
        }




    }

}
