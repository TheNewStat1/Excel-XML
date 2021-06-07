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
    class ExcelStyleSheet
    {
        public void GenerateWorkbookStylesPart(ref WorkbookStylesPart part)
        {
            // Init XML Markup
            MarkupCompatibilityAttributes markupCompatibilityAttributes1 = new MarkupCompatibilityAttributes();
            markupCompatibilityAttributes1.Ignorable = "x14ac x16r2 xr";

            // Create Stylesheet File
            Stylesheet stylesheet = new Stylesheet();

            // Add Attributes to stylesheet Tag
            stylesheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            stylesheet.AddNamespaceDeclaration("x16r2", "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");
            stylesheet.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            stylesheet.MCAttributes = markupCompatibilityAttributes1;


            // Init Fonts Markup
            Fonts fonts = new Fonts();
            fonts.Count = 1u; // Specify number of fonts
            fonts.KnownFonts = true;

            // Create  Defualt Font
            Font defaultFont = new Font();

            FontSize fontSize = new FontSize();
            fontSize.Val = 11D;
            defaultFont.Append(fontSize);

            Color color = new Color();
            color.Theme = 1u;
            defaultFont.Append(color);

            FontName fontName = new FontName();
            fontName.Val = "Calibri";
            defaultFont.Append(fontName);

            FontFamilyNumbering fontFamilyNumbering = new FontFamilyNumbering();
            fontFamilyNumbering.Val = 2;
            defaultFont.Append(fontFamilyNumbering);

            FontScheme fontScheme = new FontScheme();
            fontScheme.Val = FontSchemeValues.Minor;
            defaultFont.Append(fontScheme);


            // Add Font to Fonts XML Sheet
            fonts.Append(defaultFont);


            Font boldFont = new Font();




        }
    }
}
