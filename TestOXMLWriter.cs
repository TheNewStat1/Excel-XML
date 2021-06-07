using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using A = DocumentFormat.OpenXml.Drawing;
using AP = DocumentFormat.OpenXml.ExtendedProperties;
using THM15 = DocumentFormat.OpenXml.Office2013.Theme;
using VT = DocumentFormat.OpenXml.VariantTypes;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using X15AC = DocumentFormat.OpenXml.Office2013.ExcelAc;
using OpenXmlSample;


namespace OpenXmlSample
{
    public class TestOCMLWriter
    {
        public static void RunWriter()
        {
            var workbook = SpreadsheetDocument.Create("SomeLargeFile.xlsx", SpreadsheetDocumentType.Workbook);
            List<OpenXmlAttribute> attributeList;
            OpenXmlWriter writer;

            workbook.AddWorkbookPart();
            WorksheetPart workSheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();

            writer = OpenXmlWriter.Create(workSheetPart);
            writer.WriteStartElement(new Worksheet());
            writer.WriteStartElement(new SheetData());

            for (int i = 1; i <= 10; ++i)
            {
                attributeList = new List<OpenXmlAttribute>();
                // this is the row index
                attributeList.Add(new OpenXmlAttribute("r", null, i.ToString()));

                writer.WriteStartElement(new Row(), attributeList);

                if (i % 10000 == 0)
                {
                    Console.WriteLine("The number of rows written: " + i);
                }

                for (int j = 1; j <= 10; ++j)
                {
                    attributeList = new List<OpenXmlAttribute>();
                    // this is the data type ("t"), with CellValues.String ("str")
                    attributeList.Add(new OpenXmlAttribute("t", null, "str"));

                    // it's suggested you also have the cell reference, but
                    // you'll have to calculate the correct cell reference yourself.
                    // Here's an example:
                    //attributeList.Add(new OpenXmlAttribute("r", null, "A1"));

                    writer.WriteStartElement(new Cell(), attributeList);

                    writer.WriteElement(new CellValue(string.Format("R{0}C{1}", i, j)));

                    // this is for Cell
                    writer.WriteEndElement();
                }

                // this is for Row
                writer.WriteEndElement();
            }

            // this is for SheetData
            writer.WriteEndElement();
            // this is for Worksheet
            writer.WriteEndElement();
            writer.Close();

            writer = OpenXmlWriter.Create(workbook.WorkbookPart);
            writer.WriteStartElement(new Workbook());
            writer.WriteStartElement(new Sheets());

            // you can use object initialisers like this only when the properties
            // are actual properties. SDK classes sometimes have property-like properties
            // but are actually classes. For example, the Cell class has the CellValue
            // "property" but is actually a child class internally.
            // If the properties correspond to actual XML attributes, then you're fine.
            writer.WriteElement(new Sheet()
            {
                Name = "Sheet1",
                SheetId = 1,
                Id = workbook.WorkbookPart.GetIdOfPart(workSheetPart)
            });

            writer.WriteEndElement(); // Write end for WorkSheet Element
            writer.WriteEndElement(); // Write end for WorkBook Element
            writer.Close();

            workbook.Close();
        }
    }
    
}
