using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

namespace OpenCML_Sandbox
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            CreateSpreadsheetWorkbook();
        }


        public static void CreateSpreadsheetWorkbook()
        {
            // Create Excel Workbook
            SpreadsheetDocument workbook = SpreadsheetDocument.Create("SimpleFile.xlsx", SpreadsheetDocumentType.Workbook);

            List<OpenXmlAttribute> attributeList;
            OpenXmlWriter writer;

            //workbook.AddWorkbookPart();
            WorksheetPart workSheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();

            writer = OpenXmlWriter.Create(workSheetPart);
            writer.WriteStartElement(new Worksheet());
            writer.WriteStartElement(new SheetData());


            attributeList = new List<OpenXmlAttribute>();
            // this is the row index


            writer.WriteStartElement(new Row(), attributeList);

            attributeList = new List<OpenXmlAttribute>();
            // this is the data type ("t"), with CellValues.String ("str")
            attributeList.Add(new OpenXmlAttribute("t", null, "str"));

            // it's suggested you also have the cell reference, but
            // you'll have to calculate the correct cell reference yourself.
            // Here's an example:
            //attributeList.Add(new OpenXmlAttribute("r", null, "A1"));

            writer.WriteStartElement(new Cell(), attributeList);

            writer.WriteElement(new CellValue("Foo"));

            // this is for Cell
            writer.WriteEndElement();

            // this is for Row
            writer.WriteEndElement();


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

