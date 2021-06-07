
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

    public class MakePackage
    {
        public static void Main(string[] args)
        {
            TestOCMLWriter.RunWriter();
            var comps = new ExcelSource();
            comps.CreatePackage(pathToFile: "New.xlsx");

            //comps.GenerateWorksheetPart();
        } 
    }


}