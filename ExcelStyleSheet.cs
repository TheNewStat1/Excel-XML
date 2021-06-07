using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

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


            // Init Fonts XML Markup
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

            // Add Default Font to Fonts XML Sheet
            fonts.Append(defaultFont);



            // Create Bold Font
            Font boldFont = new Font();

            Bold boldStyle = new Bold();
            boldFont.Append(boldStyle);
            boldFont.Append(fontSize);

            color = new Color();
            color.Theme = 1u;
            boldFont.Append(color);

            fontName = new FontName();
            fontName.Val = "Calibri";
            boldFont.Append(fontName);

            fontFamilyNumbering = new FontFamilyNumbering();
            fontFamilyNumbering.Val = 2;
            boldFont.Append(fontFamilyNumbering);

            fontScheme = new FontScheme();
            fontScheme.Val = FontSchemeValues.Minor;
            boldFont.Append(fontScheme);

            // Add Bold Font to Fonts XML Sheet
            fonts.Append(boldFont);

            // Add Fonts Markup to XML
            stylesheet.Append(fonts);






            Fills fills = new Fills();
            fills.Count = 2u;

            Fill fill = new Fill();

            PatternFill patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.None;

            fill.Append(patternFill);

            fills.Append(fill);

            Fill fill1 = new Fill();

            PatternFill patternFill1 = new PatternFill();
            patternFill1.PatternType = PatternValues.Gray125;

            fill1.Append(patternFill1);

            fills.Append(fill1);

            stylesheet.Append(fills);


        }
    }
}
