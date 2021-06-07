namespace OpenXmlSample
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using A = DocumentFormat.OpenXml.Drawing;
    using AP = DocumentFormat.OpenXml.ExtendedProperties;
    using THM15 = DocumentFormat.OpenXml.Office2013.Theme;
    using VT = DocumentFormat.OpenXml.VariantTypes;
    using X14 = DocumentFormat.OpenXml.Office2010.Excel;
    using X15 = DocumentFormat.OpenXml.Office2013.Excel;
    using X15AC = DocumentFormat.OpenXml.Office2013.ExcelAc;


    public class ExcelSource
    {

        public void CreatePackage(String pathToFile)
        {
            SpreadsheetDocument pkg = null;
            try
            {
                pkg = SpreadsheetDocument.Create(pathToFile, SpreadsheetDocumentType.Workbook);

                this.CreateParts(ref pkg);
            }
            finally
            {
                if ((pkg != null))
                {
                    pkg.Dispose();
                }
            }
            //return pkg;
        }

        public void CreateParts(ref SpreadsheetDocument pkg)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart = pkg.AddExtendedFilePropertiesPart();
            pkg.ChangeIdOfPart(extendedFilePropertiesPart, "rId3");
            this.GenerateExtendedFilePropertiesPart(ref extendedFilePropertiesPart);

            CoreFilePropertiesPart coreFilePropertiesPart = pkg.AddCoreFilePropertiesPart();
            pkg.ChangeIdOfPart(coreFilePropertiesPart, "rId2");
            this.GenerateCoreFilePropertiesPart(ref coreFilePropertiesPart);

            WorkbookPart workbookPart = pkg.AddWorkbookPart();
            this.GenerateWorkbookPart(ref workbookPart);

            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>("rId3");
            this.GenerateWorkbookStylesPart(ref workbookStylesPart);

            ThemePart themePart = workbookPart.AddNewPart<ThemePart>("rId2");
            this.GenerateThemePart(ref themePart);

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
            this.GenerateWorksheetPart(ref worksheetPart);

            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart = worksheetPart.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            this.GenerateSpreadsheetPrinterSettingsPart(ref spreadsheetPrinterSettingsPart);

            SharedStringTablePart sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>("rId4");
            this.GenerateSharedStringTablePart(ref sharedStringTablePart);

        }

        public void GenerateExtendedFilePropertiesPart(ref ExtendedFilePropertiesPart part)
        {
            AP.Properties apProperties = new AP.Properties();

            apProperties.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");

            AP.Application apApplication = new AP.Application("Microsoft Excel");

            apProperties.Append(apApplication);

            AP.DocumentSecurity apDocumentSecurity = new AP.DocumentSecurity("0");

            apProperties.Append(apDocumentSecurity);

            AP.ScaleCrop apScaleCrop = new AP.ScaleCrop("false");

            apProperties.Append(apScaleCrop);

            AP.HeadingPairs apHeadingPairs = new AP.HeadingPairs();

            VT.VTVector vtVTVector = new VT.VTVector();
            vtVTVector.Size = 2u;
            vtVTVector.BaseType = VT.VectorBaseValues.Variant;

            VT.Variant vtVariant = new VT.Variant();

            VT.VTLPSTR vtVTLPSTR = new VT.VTLPSTR("Worksheets");

            vtVariant.Append(vtVTLPSTR);

            vtVTVector.Append(vtVariant);

            VT.Variant vtVariant1 = new VT.Variant();

            VT.VTInt32 vtVTInt32 = new VT.VTInt32("1");

            vtVariant1.Append(vtVTInt32);

            vtVTVector.Append(vtVariant1);

            apHeadingPairs.Append(vtVTVector);

            apProperties.Append(apHeadingPairs);

            AP.TitlesOfParts apTitlesOfParts = new AP.TitlesOfParts();

            VT.VTVector vtVTVector1 = new VT.VTVector();
            vtVTVector1.Size = 1u;
            vtVTVector1.BaseType = VT.VectorBaseValues.Lpstr;

            VT.VTLPSTR vtVTLPSTR1 = new VT.VTLPSTR("Sheet1");

            vtVTVector1.Append(vtVTLPSTR1);

            apTitlesOfParts.Append(vtVTVector1);

            apProperties.Append(apTitlesOfParts);

            AP.Company apCompany = new AP.Company("");

            apProperties.Append(apCompany);

            AP.LinksUpToDate apLinksUpToDate = new AP.LinksUpToDate("false");

            apProperties.Append(apLinksUpToDate);

            AP.SharedDocument apSharedDocument = new AP.SharedDocument("false");

            apProperties.Append(apSharedDocument);

            AP.HyperlinksChanged apHyperlinksChanged = new AP.HyperlinksChanged("false");

            apProperties.Append(apHyperlinksChanged);

            AP.ApplicationVersion apApplicationVersion = new AP.ApplicationVersion("16.0300");

            apProperties.Append(apApplicationVersion);

            part.Properties = apProperties;
        }

        public void GenerateCoreFilePropertiesPart(ref CoreFilePropertiesPart part)
        {
            string base64 = @"PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/Pg0KPGNwOmNvcmVQcm9wZXJ0aWVzIHhtbG5zOmNwPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvcGFja2FnZS8yMDA2L21ldGFkYXRhL2NvcmUtcHJvcGVydGllcyIgeG1sbnM6ZGM9Imh0dHA6Ly9wdXJsLm9yZy9kYy9lbGVtZW50cy8xLjEvIiB4bWxuczpkY3Rlcm1zPSJodHRwOi8vcHVybC5vcmcvZGMvdGVybXMvIiB4bWxuczpkY21pdHlwZT0iaHR0cDovL3B1cmwub3JnL2RjL2RjbWl0eXBlLyIgeG1sbnM6eHNpPSJodHRwOi8vd3d3LnczLm9yZy8yMDAxL1hNTFNjaGVtYS1pbnN0YW5jZSI+PGRjOmNyZWF0b3I+UHJhdGlrIFBhdGVsPC9kYzpjcmVhdG9yPjxjcDpsYXN0TW9kaWZpZWRCeT5QcmF0aWsgUGF0ZWw8L2NwOmxhc3RNb2RpZmllZEJ5PjxkY3Rlcm1zOmNyZWF0ZWQgeHNpOnR5cGU9ImRjdGVybXM6VzNDRFRGIj4yMDIxLTA2LTA2VDA0OjQ2OjAxWjwvZGN0ZXJtczpjcmVhdGVkPjxkY3Rlcm1zOm1vZGlmaWVkIHhzaTp0eXBlPSJkY3Rlcm1zOlczQ0RURiI+MjAyMS0wNi0wNlQwNDo0OTozOFo8L2RjdGVybXM6bW9kaWZpZWQ+PC9jcDpjb3JlUHJvcGVydGllcz4=";

            Stream mem = new MemoryStream(Convert.FromBase64String(base64), false);
            try
            {
                part.FeedData(mem);
            }
            finally
            {
                mem.Dispose();
            }
        }

        public void GenerateWorkbookPart(ref WorkbookPart part)
        {
            MarkupCompatibilityAttributes markupCompatibilityAttributes = new MarkupCompatibilityAttributes();
            markupCompatibilityAttributes.Ignorable = "x15 xr xr6 xr10 xr2";

            Workbook workbook = new Workbook();

            workbook.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            workbook.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            workbook.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            workbook.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            workbook.AddNamespaceDeclaration("xr6", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6");
            workbook.AddNamespaceDeclaration("xr10", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10");
            workbook.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");

            workbook.MCAttributes = markupCompatibilityAttributes;

            FileVersion fileVersion = new FileVersion();
            fileVersion.ApplicationName = "xl";
            fileVersion.LastEdited = "7";
            fileVersion.LowestEdited = "7";
            fileVersion.BuildVersion = "24026";

            workbook.Append(fileVersion);

            WorkbookProperties workbookProperties = new WorkbookProperties();
            workbookProperties.DefaultThemeVersion = 166925u;

            workbook.Append(workbookProperties);

            AlternateContent alternateContent = new AlternateContent();

            alternateContent.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

            AlternateContentChoice alternateContentChoice = new AlternateContentChoice();
            alternateContentChoice.Requires = "x15";

            X15AC.AbsolutePath x15acAbsolutePath = new X15AC.AbsolutePath();

            x15acAbsolutePath.AddNamespaceDeclaration("x15ac", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac");

            x15acAbsolutePath.Url = "C:\\Users\\pprat\\Desktop\\";

            alternateContentChoice.Append(x15acAbsolutePath);

            alternateContent.Append(alternateContentChoice);

            workbook.Append(alternateContent);

            //OpenXmlUnknownElement openXmlUnknownElement = new OpenXmlUnknownElement("xr", "revisionPtr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            //workbook.Append(openXmlUnknownElement);

            BookViews bookViews = new BookViews();

            WorkbookView workbookView = new WorkbookView();
            workbookView.XWindow = -120;
            workbookView.YWindow = -120;
            workbookView.WindowWidth = 20730u;
            workbookView.WindowHeight = 11160u;

            bookViews.Append(workbookView);

            workbook.Append(bookViews);

            Sheets sheets = new Sheets();

            Sheet sheet = new Sheet();
            sheet.Name = "Sheet1";
            sheet.SheetId = 1u;
            sheet.Id = "rId1";

            sheets.Append(sheet);

            workbook.Append(sheets);

            CalculationProperties calculationProperties = new CalculationProperties();
            calculationProperties.CalculationId = 191029u;

            workbook.Append(calculationProperties);

            WorkbookExtensionList workbookExtensionList = new WorkbookExtensionList();

            WorkbookExtension workbookExtension = new WorkbookExtension();

            workbookExtension.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");

            workbookExtension.Uri = "{140A7094-0E35-4892-8432-C4D2E57EDEB5}";

            X15.WorkbookProperties x15WorkbookProperties = new X15.WorkbookProperties();
            x15WorkbookProperties.ChartTrackingReferenceBase = true;

            workbookExtension.Append(x15WorkbookProperties);

            workbookExtensionList.Append(workbookExtension);

            WorkbookExtension workbookExtension1 = new WorkbookExtension();

            workbookExtension1.Uri = "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}";

            workbookExtensionList.Append(workbookExtension1);

            workbook.Append(workbookExtensionList);

            part.Workbook = workbook;
        }

        public void GenerateWorkbookStylesPart(ref WorkbookStylesPart part)
        {
            MarkupCompatibilityAttributes markupCompatibilityAttributes1 = new MarkupCompatibilityAttributes();
            markupCompatibilityAttributes1.Ignorable = "x14ac x16r2 xr";

            Stylesheet stylesheet = new Stylesheet();

            stylesheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            stylesheet.AddNamespaceDeclaration("x16r2", "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");
            stylesheet.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");

            stylesheet.MCAttributes = markupCompatibilityAttributes1;



            // Create Fonts Used for Excel file
            Fonts fonts = new Fonts();
            fonts.Count = 3u; // Specify Number if Fonts
            fonts.KnownFonts = true;



            Font font = new Font();

            FontSize fontSize = new FontSize();
            fontSize.Val = 11D;

            font.Append(fontSize);

            Color color = new Color();
            color.Theme = 1u;

            font.Append(color);

            FontName fontName = new FontName();
            fontName.Val = "Calibri";

            font.Append(fontName);

            FontFamilyNumbering fontFamilyNumbering = new FontFamilyNumbering();
            fontFamilyNumbering.Val = 2;

            font.Append(fontFamilyNumbering);

            FontScheme fontScheme = new FontScheme();
            fontScheme.Val = FontSchemeValues.Minor;

            font.Append(fontScheme);

            fonts.Append(font);






            Font font1 = new Font();

            Bold bold = new Bold();

            font1.Append(bold);

            FontSize fontSize1 = new FontSize();
            fontSize1.Val = 11D;

            font1.Append(fontSize1);

            Color color1 = new Color();
            color1.Theme = 1u;

            font1.Append(color1);

            FontName fontName1 = new FontName();
            fontName1.Val = "Calibri";

            font1.Append(fontName1);

            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering();
            fontFamilyNumbering1.Val = 2;

            font1.Append(fontFamilyNumbering1);

            FontScheme fontScheme1 = new FontScheme();
            fontScheme1.Val = FontSchemeValues.Minor;

            font1.Append(fontScheme1);

            fonts.Append(font1);

            Font font2 = new Font();

            FontSize fontSize2 = new FontSize();
            fontSize2.Val = 8D;

            font2.Append(fontSize2);

            FontName fontName2 = new FontName();
            fontName2.Val = "Calibri";

            font2.Append(fontName2);

            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering();
            fontFamilyNumbering2.Val = 2;

            font2.Append(fontFamilyNumbering2);

            FontScheme fontScheme2 = new FontScheme();
            fontScheme2.Val = FontSchemeValues.Minor;

            font2.Append(fontScheme2);

            fonts.Append(font2);

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

            Borders borders = new Borders();
            borders.Count = 1u;

            Border border = new Border();

            LeftBorder leftBorder = new LeftBorder();

            border.Append(leftBorder);

            RightBorder rightBorder = new RightBorder();

            border.Append(rightBorder);

            TopBorder topBorder = new TopBorder();

            border.Append(topBorder);

            BottomBorder bottomBorder = new BottomBorder();

            border.Append(bottomBorder);

            DiagonalBorder diagonalBorder = new DiagonalBorder();

            border.Append(diagonalBorder);

            borders.Append(border);

            stylesheet.Append(borders);

            CellStyleFormats cellStyleFormats = new CellStyleFormats();
            cellStyleFormats.Count = 1u;

            CellFormat cellFormat = new CellFormat();
            cellFormat.NumberFormatId = 0u;
            cellFormat.FontId = 0u;
            cellFormat.FillId = 0u;
            cellFormat.BorderId = 0u;

            cellStyleFormats.Append(cellFormat);

            stylesheet.Append(cellStyleFormats);

            CellFormats cellFormats = new CellFormats();
            cellFormats.Count = 7u;

            CellFormat cellFormat1 = new CellFormat();
            cellFormat1.NumberFormatId = 0u;
            cellFormat1.FontId = 0u;
            cellFormat1.FillId = 0u;
            cellFormat1.BorderId = 0u;
            cellFormat1.FormatId = 0u;

            cellFormats.Append(cellFormat1);

            CellFormat cellFormat2 = new CellFormat();
            cellFormat2.NumberFormatId = 0u;
            cellFormat2.FontId = 1u;
            cellFormat2.FillId = 0u;
            cellFormat2.BorderId = 0u;
            cellFormat2.FormatId = 0u;
            cellFormat2.ApplyFont = true;
            cellFormat2.ApplyAlignment = true;

            Alignment alignment = new Alignment();
            alignment.Horizontal = HorizontalAlignmentValues.Center;

            cellFormat2.Append(alignment);

            cellFormats.Append(cellFormat2);

            CellFormat cellFormat3 = new CellFormat();
            cellFormat3.NumberFormatId = 0u;
            cellFormat3.FontId = 0u;
            cellFormat3.FillId = 0u;
            cellFormat3.BorderId = 0u;
            cellFormat3.FormatId = 0u;
            cellFormat3.ApplyNumberFormat = true;
            cellFormat3.ApplyAlignment = true;

            Alignment alignment1 = new Alignment();
            alignment1.Horizontal = HorizontalAlignmentValues.Center;

            cellFormat3.Append(alignment1);

            cellFormats.Append(cellFormat3);

            CellFormat cellFormat4 = new CellFormat();
            cellFormat4.NumberFormatId = 49u;
            cellFormat4.FontId = 0u;
            cellFormat4.FillId = 0u;
            cellFormat4.BorderId = 0u;
            cellFormat4.FormatId = 0u;
            cellFormat4.QuotePrefix = true;
            cellFormat4.ApplyNumberFormat = true;
            cellFormat4.ApplyAlignment = true;

            Alignment alignment2 = new Alignment();
            alignment2.Horizontal = HorizontalAlignmentValues.Center;

            cellFormat4.Append(alignment2);

            cellFormats.Append(cellFormat4);

            CellFormat cellFormat5 = new CellFormat();
            cellFormat5.NumberFormatId = 0u;
            cellFormat5.FontId = 0u;
            cellFormat5.FillId = 0u;
            cellFormat5.BorderId = 0u;
            cellFormat5.FormatId = 0u;
            cellFormat5.ApplyAlignment = true;

            Alignment alignment3 = new Alignment();
            alignment3.Horizontal = HorizontalAlignmentValues.Center;

            cellFormat5.Append(alignment3);

            cellFormats.Append(cellFormat5);

            CellFormat cellFormat6 = new CellFormat();
            cellFormat6.NumberFormatId = 14u;
            cellFormat6.FontId = 0u;
            cellFormat6.FillId = 0u;
            cellFormat6.BorderId = 0u;
            cellFormat6.FormatId = 0u;
            cellFormat6.ApplyNumberFormat = true;
            cellFormat6.ApplyAlignment = true;

            Alignment alignment4 = new Alignment();
            alignment4.Horizontal = HorizontalAlignmentValues.Center;

            cellFormat6.Append(alignment4);

            cellFormats.Append(cellFormat6);

            CellFormat cellFormat7 = new CellFormat();
            cellFormat7.NumberFormatId = 49u;
            cellFormat7.FontId = 0u;
            cellFormat7.FillId = 0u;
            cellFormat7.BorderId = 0u;
            cellFormat7.FormatId = 0u;
            cellFormat7.ApplyNumberFormat = true;
            cellFormat7.ApplyAlignment = true;

            Alignment alignment5 = new Alignment();
            alignment5.Horizontal = HorizontalAlignmentValues.Center;

            cellFormat7.Append(alignment5);

            cellFormats.Append(cellFormat7);

            stylesheet.Append(cellFormats);

            CellStyles cellStyles = new CellStyles();
            cellStyles.Count = 1u;

            CellStyle cellStyle = new CellStyle();
            cellStyle.Name = "Normal";
            cellStyle.FormatId = 0u;
            cellStyle.BuiltinId = 0u;

            cellStyles.Append(cellStyle);

            stylesheet.Append(cellStyles);

            DifferentialFormats differentialFormats = new DifferentialFormats();
            differentialFormats.Count = 0u;

            stylesheet.Append(differentialFormats);

            TableStyles tableStyles = new TableStyles();
            tableStyles.Count = 0u;
            tableStyles.DefaultTableStyle = "TableStyleMedium2";
            tableStyles.DefaultPivotStyle = "PivotStyleLight16";

            stylesheet.Append(tableStyles);

            StylesheetExtensionList stylesheetExtensionList = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension = new StylesheetExtension();

            stylesheetExtension.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");

            stylesheetExtension.Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}";

            X14.SlicerStyles x14SlicerStyles = new X14.SlicerStyles();
            x14SlicerStyles.DefaultSlicerStyle = "SlicerStyleLight1";

            stylesheetExtension.Append(x14SlicerStyles);

            stylesheetExtensionList.Append(stylesheetExtension);

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension();

            stylesheetExtension1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");

            stylesheetExtension1.Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}";

            X15.TimelineStyles x15TimelineStyles = new X15.TimelineStyles();
            x15TimelineStyles.DefaultTimelineStyle = "TimeSlicerStyleLight1";

            stylesheetExtension1.Append(x15TimelineStyles);

            stylesheetExtensionList.Append(stylesheetExtension1);

            stylesheet.Append(stylesheetExtensionList);

            part.Stylesheet = stylesheet;
        }

        public void GenerateThemePart(ref ThemePart part)
        {
            A.Theme aTheme = new A.Theme();

            aTheme.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            aTheme.Name = "Office Theme";

            A.ThemeElements aThemeElements = new A.ThemeElements();

            A.ColorScheme aColorScheme = new A.ColorScheme();
            aColorScheme.Name = "Office";

            A.Dark1Color aDark1Color = new A.Dark1Color();

            A.SystemColor aSystemColor = new A.SystemColor();
            aSystemColor.LastColor = "000000";
            aSystemColor.Val = A.SystemColorValues.WindowText;

            aDark1Color.Append(aSystemColor);

            aColorScheme.Append(aDark1Color);

            A.Light1Color aLight1Color = new A.Light1Color();

            A.SystemColor aSystemColor1 = new A.SystemColor();
            aSystemColor1.LastColor = "FFFFFF";
            aSystemColor1.Val = A.SystemColorValues.Window;

            aLight1Color.Append(aSystemColor1);

            aColorScheme.Append(aLight1Color);

            A.Dark2Color aDark2Color = new A.Dark2Color();

            A.RgbColorModelHex aRgbColorModelHex = new A.RgbColorModelHex();
            aRgbColorModelHex.Val = "44546A";

            aDark2Color.Append(aRgbColorModelHex);

            aColorScheme.Append(aDark2Color);

            A.Light2Color aLight2Color = new A.Light2Color();

            A.RgbColorModelHex aRgbColorModelHex1 = new A.RgbColorModelHex();
            aRgbColorModelHex1.Val = "E7E6E6";

            aLight2Color.Append(aRgbColorModelHex1);

            aColorScheme.Append(aLight2Color);

            A.Accent1Color aAccent1Color = new A.Accent1Color();

            A.RgbColorModelHex aRgbColorModelHex2 = new A.RgbColorModelHex();
            aRgbColorModelHex2.Val = "4472C4";

            aAccent1Color.Append(aRgbColorModelHex2);

            aColorScheme.Append(aAccent1Color);

            A.Accent2Color aAccent2Color = new A.Accent2Color();

            A.RgbColorModelHex aRgbColorModelHex3 = new A.RgbColorModelHex();
            aRgbColorModelHex3.Val = "ED7D31";

            aAccent2Color.Append(aRgbColorModelHex3);

            aColorScheme.Append(aAccent2Color);

            A.Accent3Color aAccent3Color = new A.Accent3Color();

            A.RgbColorModelHex aRgbColorModelHex4 = new A.RgbColorModelHex();
            aRgbColorModelHex4.Val = "A5A5A5";

            aAccent3Color.Append(aRgbColorModelHex4);

            aColorScheme.Append(aAccent3Color);

            A.Accent4Color aAccent4Color = new A.Accent4Color();

            A.RgbColorModelHex aRgbColorModelHex5 = new A.RgbColorModelHex();
            aRgbColorModelHex5.Val = "FFC000";

            aAccent4Color.Append(aRgbColorModelHex5);

            aColorScheme.Append(aAccent4Color);

            A.Accent5Color aAccent5Color = new A.Accent5Color();

            A.RgbColorModelHex aRgbColorModelHex6 = new A.RgbColorModelHex();
            aRgbColorModelHex6.Val = "5B9BD5";

            aAccent5Color.Append(aRgbColorModelHex6);

            aColorScheme.Append(aAccent5Color);

            A.Accent6Color aAccent6Color = new A.Accent6Color();

            A.RgbColorModelHex aRgbColorModelHex7 = new A.RgbColorModelHex();
            aRgbColorModelHex7.Val = "70AD47";

            aAccent6Color.Append(aRgbColorModelHex7);

            aColorScheme.Append(aAccent6Color);

            A.Hyperlink aHyperlink = new A.Hyperlink();

            A.RgbColorModelHex aRgbColorModelHex8 = new A.RgbColorModelHex();
            aRgbColorModelHex8.Val = "0563C1";

            aHyperlink.Append(aRgbColorModelHex8);

            aColorScheme.Append(aHyperlink);

            A.FollowedHyperlinkColor aFollowedHyperlinkColor = new A.FollowedHyperlinkColor();

            A.RgbColorModelHex aRgbColorModelHex9 = new A.RgbColorModelHex();
            aRgbColorModelHex9.Val = "954F72";

            aFollowedHyperlinkColor.Append(aRgbColorModelHex9);

            aColorScheme.Append(aFollowedHyperlinkColor);

            aThemeElements.Append(aColorScheme);

            A.FontScheme aFontScheme = new A.FontScheme();
            aFontScheme.Name = "Office";

            A.MajorFont aMajorFont = new A.MajorFont();

            A.LatinFont aLatinFont = new A.LatinFont();
            aLatinFont.Typeface = "Calibri Light";
            aLatinFont.Panose = "020F0302020204030204";

            aMajorFont.Append(aLatinFont);

            A.EastAsianFont aEastAsianFont = new A.EastAsianFont();
            aEastAsianFont.Typeface = "";

            aMajorFont.Append(aEastAsianFont);

            A.ComplexScriptFont aComplexScriptFont = new A.ComplexScriptFont();
            aComplexScriptFont.Typeface = "";

            aMajorFont.Append(aComplexScriptFont);

            A.SupplementalFont aSupplementalFont = new A.SupplementalFont();
            aSupplementalFont.Script = "Jpan";
            aSupplementalFont.Typeface = "游ゴシック Light";

            aMajorFont.Append(aSupplementalFont);

            A.SupplementalFont aSupplementalFont1 = new A.SupplementalFont();
            aSupplementalFont1.Script = "Hang";
            aSupplementalFont1.Typeface = "맑은 고딕";

            aMajorFont.Append(aSupplementalFont1);

            A.SupplementalFont aSupplementalFont2 = new A.SupplementalFont();
            aSupplementalFont2.Script = "Hans";
            aSupplementalFont2.Typeface = "等线 Light";

            aMajorFont.Append(aSupplementalFont2);

            A.SupplementalFont aSupplementalFont3 = new A.SupplementalFont();
            aSupplementalFont3.Script = "Hant";
            aSupplementalFont3.Typeface = "新細明體";

            aMajorFont.Append(aSupplementalFont3);

            A.SupplementalFont aSupplementalFont4 = new A.SupplementalFont();
            aSupplementalFont4.Script = "Arab";
            aSupplementalFont4.Typeface = "Times New Roman";

            aMajorFont.Append(aSupplementalFont4);

            A.SupplementalFont aSupplementalFont5 = new A.SupplementalFont();
            aSupplementalFont5.Script = "Hebr";
            aSupplementalFont5.Typeface = "Times New Roman";

            aMajorFont.Append(aSupplementalFont5);

            A.SupplementalFont aSupplementalFont6 = new A.SupplementalFont();
            aSupplementalFont6.Script = "Thai";
            aSupplementalFont6.Typeface = "Tahoma";

            aMajorFont.Append(aSupplementalFont6);

            A.SupplementalFont aSupplementalFont7 = new A.SupplementalFont();
            aSupplementalFont7.Script = "Ethi";
            aSupplementalFont7.Typeface = "Nyala";

            aMajorFont.Append(aSupplementalFont7);

            A.SupplementalFont aSupplementalFont8 = new A.SupplementalFont();
            aSupplementalFont8.Script = "Beng";
            aSupplementalFont8.Typeface = "Vrinda";

            aMajorFont.Append(aSupplementalFont8);

            A.SupplementalFont aSupplementalFont9 = new A.SupplementalFont();
            aSupplementalFont9.Script = "Gujr";
            aSupplementalFont9.Typeface = "Shruti";

            aMajorFont.Append(aSupplementalFont9);

            A.SupplementalFont aSupplementalFont10 = new A.SupplementalFont();
            aSupplementalFont10.Script = "Khmr";
            aSupplementalFont10.Typeface = "MoolBoran";

            aMajorFont.Append(aSupplementalFont10);

            A.SupplementalFont aSupplementalFont11 = new A.SupplementalFont();
            aSupplementalFont11.Script = "Knda";
            aSupplementalFont11.Typeface = "Tunga";

            aMajorFont.Append(aSupplementalFont11);

            A.SupplementalFont aSupplementalFont12 = new A.SupplementalFont();
            aSupplementalFont12.Script = "Guru";
            aSupplementalFont12.Typeface = "Raavi";

            aMajorFont.Append(aSupplementalFont12);

            A.SupplementalFont aSupplementalFont13 = new A.SupplementalFont();
            aSupplementalFont13.Script = "Cans";
            aSupplementalFont13.Typeface = "Euphemia";

            aMajorFont.Append(aSupplementalFont13);

            A.SupplementalFont aSupplementalFont14 = new A.SupplementalFont();
            aSupplementalFont14.Script = "Cher";
            aSupplementalFont14.Typeface = "Plantagenet Cherokee";

            aMajorFont.Append(aSupplementalFont14);

            A.SupplementalFont aSupplementalFont15 = new A.SupplementalFont();
            aSupplementalFont15.Script = "Yiii";
            aSupplementalFont15.Typeface = "Microsoft Yi Baiti";

            aMajorFont.Append(aSupplementalFont15);

            A.SupplementalFont aSupplementalFont16 = new A.SupplementalFont();
            aSupplementalFont16.Script = "Tibt";
            aSupplementalFont16.Typeface = "Microsoft Himalaya";

            aMajorFont.Append(aSupplementalFont16);

            A.SupplementalFont aSupplementalFont17 = new A.SupplementalFont();
            aSupplementalFont17.Script = "Thaa";
            aSupplementalFont17.Typeface = "MV Boli";

            aMajorFont.Append(aSupplementalFont17);

            A.SupplementalFont aSupplementalFont18 = new A.SupplementalFont();
            aSupplementalFont18.Script = "Deva";
            aSupplementalFont18.Typeface = "Mangal";

            aMajorFont.Append(aSupplementalFont18);

            A.SupplementalFont aSupplementalFont19 = new A.SupplementalFont();
            aSupplementalFont19.Script = "Telu";
            aSupplementalFont19.Typeface = "Gautami";

            aMajorFont.Append(aSupplementalFont19);

            A.SupplementalFont aSupplementalFont20 = new A.SupplementalFont();
            aSupplementalFont20.Script = "Taml";
            aSupplementalFont20.Typeface = "Latha";

            aMajorFont.Append(aSupplementalFont20);

            A.SupplementalFont aSupplementalFont21 = new A.SupplementalFont();
            aSupplementalFont21.Script = "Syrc";
            aSupplementalFont21.Typeface = "Estrangelo Edessa";

            aMajorFont.Append(aSupplementalFont21);

            A.SupplementalFont aSupplementalFont22 = new A.SupplementalFont();
            aSupplementalFont22.Script = "Orya";
            aSupplementalFont22.Typeface = "Kalinga";

            aMajorFont.Append(aSupplementalFont22);

            A.SupplementalFont aSupplementalFont23 = new A.SupplementalFont();
            aSupplementalFont23.Script = "Mlym";
            aSupplementalFont23.Typeface = "Kartika";

            aMajorFont.Append(aSupplementalFont23);

            A.SupplementalFont aSupplementalFont24 = new A.SupplementalFont();
            aSupplementalFont24.Script = "Laoo";
            aSupplementalFont24.Typeface = "DokChampa";

            aMajorFont.Append(aSupplementalFont24);

            A.SupplementalFont aSupplementalFont25 = new A.SupplementalFont();
            aSupplementalFont25.Script = "Sinh";
            aSupplementalFont25.Typeface = "Iskoola Pota";

            aMajorFont.Append(aSupplementalFont25);

            A.SupplementalFont aSupplementalFont26 = new A.SupplementalFont();
            aSupplementalFont26.Script = "Mong";
            aSupplementalFont26.Typeface = "Mongolian Baiti";

            aMajorFont.Append(aSupplementalFont26);

            A.SupplementalFont aSupplementalFont27 = new A.SupplementalFont();
            aSupplementalFont27.Script = "Viet";
            aSupplementalFont27.Typeface = "Times New Roman";

            aMajorFont.Append(aSupplementalFont27);

            A.SupplementalFont aSupplementalFont28 = new A.SupplementalFont();
            aSupplementalFont28.Script = "Uigh";
            aSupplementalFont28.Typeface = "Microsoft Uighur";

            aMajorFont.Append(aSupplementalFont28);

            A.SupplementalFont aSupplementalFont29 = new A.SupplementalFont();
            aSupplementalFont29.Script = "Geor";
            aSupplementalFont29.Typeface = "Sylfaen";

            aMajorFont.Append(aSupplementalFont29);

            A.SupplementalFont aSupplementalFont30 = new A.SupplementalFont();
            aSupplementalFont30.Script = "Armn";
            aSupplementalFont30.Typeface = "Arial";

            aMajorFont.Append(aSupplementalFont30);

            A.SupplementalFont aSupplementalFont31 = new A.SupplementalFont();
            aSupplementalFont31.Script = "Bugi";
            aSupplementalFont31.Typeface = "Leelawadee UI";

            aMajorFont.Append(aSupplementalFont31);

            A.SupplementalFont aSupplementalFont32 = new A.SupplementalFont();
            aSupplementalFont32.Script = "Bopo";
            aSupplementalFont32.Typeface = "Microsoft JhengHei";

            aMajorFont.Append(aSupplementalFont32);

            A.SupplementalFont aSupplementalFont33 = new A.SupplementalFont();
            aSupplementalFont33.Script = "Java";
            aSupplementalFont33.Typeface = "Javanese Text";

            aMajorFont.Append(aSupplementalFont33);

            A.SupplementalFont aSupplementalFont34 = new A.SupplementalFont();
            aSupplementalFont34.Script = "Lisu";
            aSupplementalFont34.Typeface = "Segoe UI";

            aMajorFont.Append(aSupplementalFont34);

            A.SupplementalFont aSupplementalFont35 = new A.SupplementalFont();
            aSupplementalFont35.Script = "Mymr";
            aSupplementalFont35.Typeface = "Myanmar Text";

            aMajorFont.Append(aSupplementalFont35);

            A.SupplementalFont aSupplementalFont36 = new A.SupplementalFont();
            aSupplementalFont36.Script = "Nkoo";
            aSupplementalFont36.Typeface = "Ebrima";

            aMajorFont.Append(aSupplementalFont36);

            A.SupplementalFont aSupplementalFont37 = new A.SupplementalFont();
            aSupplementalFont37.Script = "Olck";
            aSupplementalFont37.Typeface = "Nirmala UI";

            aMajorFont.Append(aSupplementalFont37);

            A.SupplementalFont aSupplementalFont38 = new A.SupplementalFont();
            aSupplementalFont38.Script = "Osma";
            aSupplementalFont38.Typeface = "Ebrima";

            aMajorFont.Append(aSupplementalFont38);

            A.SupplementalFont aSupplementalFont39 = new A.SupplementalFont();
            aSupplementalFont39.Script = "Phag";
            aSupplementalFont39.Typeface = "Phagspa";

            aMajorFont.Append(aSupplementalFont39);

            A.SupplementalFont aSupplementalFont40 = new A.SupplementalFont();
            aSupplementalFont40.Script = "Syrn";
            aSupplementalFont40.Typeface = "Estrangelo Edessa";

            aMajorFont.Append(aSupplementalFont40);

            A.SupplementalFont aSupplementalFont41 = new A.SupplementalFont();
            aSupplementalFont41.Script = "Syrj";
            aSupplementalFont41.Typeface = "Estrangelo Edessa";

            aMajorFont.Append(aSupplementalFont41);

            A.SupplementalFont aSupplementalFont42 = new A.SupplementalFont();
            aSupplementalFont42.Script = "Syre";
            aSupplementalFont42.Typeface = "Estrangelo Edessa";

            aMajorFont.Append(aSupplementalFont42);

            A.SupplementalFont aSupplementalFont43 = new A.SupplementalFont();
            aSupplementalFont43.Script = "Sora";
            aSupplementalFont43.Typeface = "Nirmala UI";

            aMajorFont.Append(aSupplementalFont43);

            A.SupplementalFont aSupplementalFont44 = new A.SupplementalFont();
            aSupplementalFont44.Script = "Tale";
            aSupplementalFont44.Typeface = "Microsoft Tai Le";

            aMajorFont.Append(aSupplementalFont44);

            A.SupplementalFont aSupplementalFont45 = new A.SupplementalFont();
            aSupplementalFont45.Script = "Talu";
            aSupplementalFont45.Typeface = "Microsoft New Tai Lue";

            aMajorFont.Append(aSupplementalFont45);

            A.SupplementalFont aSupplementalFont46 = new A.SupplementalFont();
            aSupplementalFont46.Script = "Tfng";
            aSupplementalFont46.Typeface = "Ebrima";

            aMajorFont.Append(aSupplementalFont46);

            aFontScheme.Append(aMajorFont);

            A.MinorFont aMinorFont = new A.MinorFont();

            A.LatinFont aLatinFont1 = new A.LatinFont();
            aLatinFont1.Typeface = "Calibri";
            aLatinFont1.Panose = "020F0502020204030204";

            aMinorFont.Append(aLatinFont1);

            A.EastAsianFont aEastAsianFont1 = new A.EastAsianFont();
            aEastAsianFont1.Typeface = "";

            aMinorFont.Append(aEastAsianFont1);

            A.ComplexScriptFont aComplexScriptFont1 = new A.ComplexScriptFont();
            aComplexScriptFont1.Typeface = "";

            aMinorFont.Append(aComplexScriptFont1);

            A.SupplementalFont aSupplementalFont47 = new A.SupplementalFont();
            aSupplementalFont47.Script = "Jpan";
            aSupplementalFont47.Typeface = "游ゴシック";

            aMinorFont.Append(aSupplementalFont47);

            A.SupplementalFont aSupplementalFont48 = new A.SupplementalFont();
            aSupplementalFont48.Script = "Hang";
            aSupplementalFont48.Typeface = "맑은 고딕";

            aMinorFont.Append(aSupplementalFont48);

            A.SupplementalFont aSupplementalFont49 = new A.SupplementalFont();
            aSupplementalFont49.Script = "Hans";
            aSupplementalFont49.Typeface = "等线";

            aMinorFont.Append(aSupplementalFont49);

            A.SupplementalFont aSupplementalFont50 = new A.SupplementalFont();
            aSupplementalFont50.Script = "Hant";
            aSupplementalFont50.Typeface = "新細明體";

            aMinorFont.Append(aSupplementalFont50);

            A.SupplementalFont aSupplementalFont51 = new A.SupplementalFont();
            aSupplementalFont51.Script = "Arab";
            aSupplementalFont51.Typeface = "Arial";

            aMinorFont.Append(aSupplementalFont51);

            A.SupplementalFont aSupplementalFont52 = new A.SupplementalFont();
            aSupplementalFont52.Script = "Hebr";
            aSupplementalFont52.Typeface = "Arial";

            aMinorFont.Append(aSupplementalFont52);

            A.SupplementalFont aSupplementalFont53 = new A.SupplementalFont();
            aSupplementalFont53.Script = "Thai";
            aSupplementalFont53.Typeface = "Tahoma";

            aMinorFont.Append(aSupplementalFont53);

            A.SupplementalFont aSupplementalFont54 = new A.SupplementalFont();
            aSupplementalFont54.Script = "Ethi";
            aSupplementalFont54.Typeface = "Nyala";

            aMinorFont.Append(aSupplementalFont54);

            A.SupplementalFont aSupplementalFont55 = new A.SupplementalFont();
            aSupplementalFont55.Script = "Beng";
            aSupplementalFont55.Typeface = "Vrinda";

            aMinorFont.Append(aSupplementalFont55);

            A.SupplementalFont aSupplementalFont56 = new A.SupplementalFont();
            aSupplementalFont56.Script = "Gujr";
            aSupplementalFont56.Typeface = "Shruti";

            aMinorFont.Append(aSupplementalFont56);

            A.SupplementalFont aSupplementalFont57 = new A.SupplementalFont();
            aSupplementalFont57.Script = "Khmr";
            aSupplementalFont57.Typeface = "DaunPenh";

            aMinorFont.Append(aSupplementalFont57);

            A.SupplementalFont aSupplementalFont58 = new A.SupplementalFont();
            aSupplementalFont58.Script = "Knda";
            aSupplementalFont58.Typeface = "Tunga";

            aMinorFont.Append(aSupplementalFont58);

            A.SupplementalFont aSupplementalFont59 = new A.SupplementalFont();
            aSupplementalFont59.Script = "Guru";
            aSupplementalFont59.Typeface = "Raavi";

            aMinorFont.Append(aSupplementalFont59);

            A.SupplementalFont aSupplementalFont60 = new A.SupplementalFont();
            aSupplementalFont60.Script = "Cans";
            aSupplementalFont60.Typeface = "Euphemia";

            aMinorFont.Append(aSupplementalFont60);

            A.SupplementalFont aSupplementalFont61 = new A.SupplementalFont();
            aSupplementalFont61.Script = "Cher";
            aSupplementalFont61.Typeface = "Plantagenet Cherokee";

            aMinorFont.Append(aSupplementalFont61);

            A.SupplementalFont aSupplementalFont62 = new A.SupplementalFont();
            aSupplementalFont62.Script = "Yiii";
            aSupplementalFont62.Typeface = "Microsoft Yi Baiti";

            aMinorFont.Append(aSupplementalFont62);

            A.SupplementalFont aSupplementalFont63 = new A.SupplementalFont();
            aSupplementalFont63.Script = "Tibt";
            aSupplementalFont63.Typeface = "Microsoft Himalaya";

            aMinorFont.Append(aSupplementalFont63);

            A.SupplementalFont aSupplementalFont64 = new A.SupplementalFont();
            aSupplementalFont64.Script = "Thaa";
            aSupplementalFont64.Typeface = "MV Boli";

            aMinorFont.Append(aSupplementalFont64);

            A.SupplementalFont aSupplementalFont65 = new A.SupplementalFont();
            aSupplementalFont65.Script = "Deva";
            aSupplementalFont65.Typeface = "Mangal";

            aMinorFont.Append(aSupplementalFont65);

            A.SupplementalFont aSupplementalFont66 = new A.SupplementalFont();
            aSupplementalFont66.Script = "Telu";
            aSupplementalFont66.Typeface = "Gautami";

            aMinorFont.Append(aSupplementalFont66);

            A.SupplementalFont aSupplementalFont67 = new A.SupplementalFont();
            aSupplementalFont67.Script = "Taml";
            aSupplementalFont67.Typeface = "Latha";

            aMinorFont.Append(aSupplementalFont67);

            A.SupplementalFont aSupplementalFont68 = new A.SupplementalFont();
            aSupplementalFont68.Script = "Syrc";
            aSupplementalFont68.Typeface = "Estrangelo Edessa";

            aMinorFont.Append(aSupplementalFont68);

            A.SupplementalFont aSupplementalFont69 = new A.SupplementalFont();
            aSupplementalFont69.Script = "Orya";
            aSupplementalFont69.Typeface = "Kalinga";

            aMinorFont.Append(aSupplementalFont69);

            A.SupplementalFont aSupplementalFont70 = new A.SupplementalFont();
            aSupplementalFont70.Script = "Mlym";
            aSupplementalFont70.Typeface = "Kartika";

            aMinorFont.Append(aSupplementalFont70);

            A.SupplementalFont aSupplementalFont71 = new A.SupplementalFont();
            aSupplementalFont71.Script = "Laoo";
            aSupplementalFont71.Typeface = "DokChampa";

            aMinorFont.Append(aSupplementalFont71);

            A.SupplementalFont aSupplementalFont72 = new A.SupplementalFont();
            aSupplementalFont72.Script = "Sinh";
            aSupplementalFont72.Typeface = "Iskoola Pota";

            aMinorFont.Append(aSupplementalFont72);

            A.SupplementalFont aSupplementalFont73 = new A.SupplementalFont();
            aSupplementalFont73.Script = "Mong";
            aSupplementalFont73.Typeface = "Mongolian Baiti";

            aMinorFont.Append(aSupplementalFont73);

            A.SupplementalFont aSupplementalFont74 = new A.SupplementalFont();
            aSupplementalFont74.Script = "Viet";
            aSupplementalFont74.Typeface = "Arial";

            aMinorFont.Append(aSupplementalFont74);

            A.SupplementalFont aSupplementalFont75 = new A.SupplementalFont();
            aSupplementalFont75.Script = "Uigh";
            aSupplementalFont75.Typeface = "Microsoft Uighur";

            aMinorFont.Append(aSupplementalFont75);

            A.SupplementalFont aSupplementalFont76 = new A.SupplementalFont();
            aSupplementalFont76.Script = "Geor";
            aSupplementalFont76.Typeface = "Sylfaen";

            aMinorFont.Append(aSupplementalFont76);

            A.SupplementalFont aSupplementalFont77 = new A.SupplementalFont();
            aSupplementalFont77.Script = "Armn";
            aSupplementalFont77.Typeface = "Arial";

            aMinorFont.Append(aSupplementalFont77);

            A.SupplementalFont aSupplementalFont78 = new A.SupplementalFont();
            aSupplementalFont78.Script = "Bugi";
            aSupplementalFont78.Typeface = "Leelawadee UI";

            aMinorFont.Append(aSupplementalFont78);

            A.SupplementalFont aSupplementalFont79 = new A.SupplementalFont();
            aSupplementalFont79.Script = "Bopo";
            aSupplementalFont79.Typeface = "Microsoft JhengHei";

            aMinorFont.Append(aSupplementalFont79);

            A.SupplementalFont aSupplementalFont80 = new A.SupplementalFont();
            aSupplementalFont80.Script = "Java";
            aSupplementalFont80.Typeface = "Javanese Text";

            aMinorFont.Append(aSupplementalFont80);

            A.SupplementalFont aSupplementalFont81 = new A.SupplementalFont();
            aSupplementalFont81.Script = "Lisu";
            aSupplementalFont81.Typeface = "Segoe UI";

            aMinorFont.Append(aSupplementalFont81);

            A.SupplementalFont aSupplementalFont82 = new A.SupplementalFont();
            aSupplementalFont82.Script = "Mymr";
            aSupplementalFont82.Typeface = "Myanmar Text";

            aMinorFont.Append(aSupplementalFont82);

            A.SupplementalFont aSupplementalFont83 = new A.SupplementalFont();
            aSupplementalFont83.Script = "Nkoo";
            aSupplementalFont83.Typeface = "Ebrima";

            aMinorFont.Append(aSupplementalFont83);

            A.SupplementalFont aSupplementalFont84 = new A.SupplementalFont();
            aSupplementalFont84.Script = "Olck";
            aSupplementalFont84.Typeface = "Nirmala UI";

            aMinorFont.Append(aSupplementalFont84);

            A.SupplementalFont aSupplementalFont85 = new A.SupplementalFont();
            aSupplementalFont85.Script = "Osma";
            aSupplementalFont85.Typeface = "Ebrima";

            aMinorFont.Append(aSupplementalFont85);

            A.SupplementalFont aSupplementalFont86 = new A.SupplementalFont();
            aSupplementalFont86.Script = "Phag";
            aSupplementalFont86.Typeface = "Phagspa";

            aMinorFont.Append(aSupplementalFont86);

            A.SupplementalFont aSupplementalFont87 = new A.SupplementalFont();
            aSupplementalFont87.Script = "Syrn";
            aSupplementalFont87.Typeface = "Estrangelo Edessa";

            aMinorFont.Append(aSupplementalFont87);

            A.SupplementalFont aSupplementalFont88 = new A.SupplementalFont();
            aSupplementalFont88.Script = "Syrj";
            aSupplementalFont88.Typeface = "Estrangelo Edessa";

            aMinorFont.Append(aSupplementalFont88);

            A.SupplementalFont aSupplementalFont89 = new A.SupplementalFont();
            aSupplementalFont89.Script = "Syre";
            aSupplementalFont89.Typeface = "Estrangelo Edessa";

            aMinorFont.Append(aSupplementalFont89);

            A.SupplementalFont aSupplementalFont90 = new A.SupplementalFont();
            aSupplementalFont90.Script = "Sora";
            aSupplementalFont90.Typeface = "Nirmala UI";

            aMinorFont.Append(aSupplementalFont90);

            A.SupplementalFont aSupplementalFont91 = new A.SupplementalFont();
            aSupplementalFont91.Script = "Tale";
            aSupplementalFont91.Typeface = "Microsoft Tai Le";

            aMinorFont.Append(aSupplementalFont91);

            A.SupplementalFont aSupplementalFont92 = new A.SupplementalFont();
            aSupplementalFont92.Script = "Talu";
            aSupplementalFont92.Typeface = "Microsoft New Tai Lue";

            aMinorFont.Append(aSupplementalFont92);

            A.SupplementalFont aSupplementalFont93 = new A.SupplementalFont();
            aSupplementalFont93.Script = "Tfng";
            aSupplementalFont93.Typeface = "Ebrima";

            aMinorFont.Append(aSupplementalFont93);

            aFontScheme.Append(aMinorFont);

            aThemeElements.Append(aFontScheme);

            A.FormatScheme aFormatScheme = new A.FormatScheme();
            aFormatScheme.Name = "Office";

            A.FillStyleList aFillStyleList = new A.FillStyleList();

            A.SolidFill aSolidFill = new A.SolidFill();

            A.SchemeColor aSchemeColor = new A.SchemeColor();
            aSchemeColor.Val = A.SchemeColorValues.PhColor;

            aSolidFill.Append(aSchemeColor);

            aFillStyleList.Append(aSolidFill);

            A.GradientFill aGradientFill = new A.GradientFill();
            aGradientFill.RotateWithShape = true;

            A.GradientStopList aGradientStopList = new A.GradientStopList();

            A.GradientStop aGradientStop = new A.GradientStop();
            aGradientStop.Position = 0;

            A.SchemeColor aSchemeColor1 = new A.SchemeColor();
            aSchemeColor1.Val = A.SchemeColorValues.PhColor;

            A.LuminanceModulation aLuminanceModulation = new A.LuminanceModulation();
            aLuminanceModulation.Val = 110000;

            aSchemeColor1.Append(aLuminanceModulation);

            A.SaturationModulation aSaturationModulation = new A.SaturationModulation();
            aSaturationModulation.Val = 105000;

            aSchemeColor1.Append(aSaturationModulation);

            A.Tint aTint = new A.Tint();
            aTint.Val = 67000;

            aSchemeColor1.Append(aTint);

            aGradientStop.Append(aSchemeColor1);

            aGradientStopList.Append(aGradientStop);

            A.GradientStop aGradientStop1 = new A.GradientStop();
            aGradientStop1.Position = 50000;

            A.SchemeColor aSchemeColor2 = new A.SchemeColor();
            aSchemeColor2.Val = A.SchemeColorValues.PhColor;

            A.LuminanceModulation aLuminanceModulation1 = new A.LuminanceModulation();
            aLuminanceModulation1.Val = 105000;

            aSchemeColor2.Append(aLuminanceModulation1);

            A.SaturationModulation aSaturationModulation1 = new A.SaturationModulation();
            aSaturationModulation1.Val = 103000;

            aSchemeColor2.Append(aSaturationModulation1);

            A.Tint aTint1 = new A.Tint();
            aTint1.Val = 73000;

            aSchemeColor2.Append(aTint1);

            aGradientStop1.Append(aSchemeColor2);

            aGradientStopList.Append(aGradientStop1);

            A.GradientStop aGradientStop2 = new A.GradientStop();
            aGradientStop2.Position = 100000;

            A.SchemeColor aSchemeColor3 = new A.SchemeColor();
            aSchemeColor3.Val = A.SchemeColorValues.PhColor;

            A.LuminanceModulation aLuminanceModulation2 = new A.LuminanceModulation();
            aLuminanceModulation2.Val = 105000;

            aSchemeColor3.Append(aLuminanceModulation2);

            A.SaturationModulation aSaturationModulation2 = new A.SaturationModulation();
            aSaturationModulation2.Val = 109000;

            aSchemeColor3.Append(aSaturationModulation2);

            A.Tint aTint2 = new A.Tint();
            aTint2.Val = 81000;

            aSchemeColor3.Append(aTint2);

            aGradientStop2.Append(aSchemeColor3);

            aGradientStopList.Append(aGradientStop2);

            aGradientFill.Append(aGradientStopList);

            A.LinearGradientFill aLinearGradientFill = new A.LinearGradientFill();
            aLinearGradientFill.Angle = 5400000;
            aLinearGradientFill.Scaled = false;

            aGradientFill.Append(aLinearGradientFill);

            aFillStyleList.Append(aGradientFill);

            A.GradientFill aGradientFill1 = new A.GradientFill();
            aGradientFill1.RotateWithShape = true;

            A.GradientStopList aGradientStopList1 = new A.GradientStopList();

            A.GradientStop aGradientStop3 = new A.GradientStop();
            aGradientStop3.Position = 0;

            A.SchemeColor aSchemeColor4 = new A.SchemeColor();
            aSchemeColor4.Val = A.SchemeColorValues.PhColor;

            A.SaturationModulation aSaturationModulation3 = new A.SaturationModulation();
            aSaturationModulation3.Val = 103000;

            aSchemeColor4.Append(aSaturationModulation3);

            A.LuminanceModulation aLuminanceModulation3 = new A.LuminanceModulation();
            aLuminanceModulation3.Val = 102000;

            aSchemeColor4.Append(aLuminanceModulation3);

            A.Tint aTint3 = new A.Tint();
            aTint3.Val = 94000;

            aSchemeColor4.Append(aTint3);

            aGradientStop3.Append(aSchemeColor4);

            aGradientStopList1.Append(aGradientStop3);

            A.GradientStop aGradientStop4 = new A.GradientStop();
            aGradientStop4.Position = 50000;

            A.SchemeColor aSchemeColor5 = new A.SchemeColor();
            aSchemeColor5.Val = A.SchemeColorValues.PhColor;

            A.SaturationModulation aSaturationModulation4 = new A.SaturationModulation();
            aSaturationModulation4.Val = 110000;

            aSchemeColor5.Append(aSaturationModulation4);

            A.LuminanceModulation aLuminanceModulation4 = new A.LuminanceModulation();
            aLuminanceModulation4.Val = 100000;

            aSchemeColor5.Append(aLuminanceModulation4);

            A.Shade aShade = new A.Shade();
            aShade.Val = 100000;

            aSchemeColor5.Append(aShade);

            aGradientStop4.Append(aSchemeColor5);

            aGradientStopList1.Append(aGradientStop4);

            A.GradientStop aGradientStop5 = new A.GradientStop();
            aGradientStop5.Position = 100000;

            A.SchemeColor aSchemeColor6 = new A.SchemeColor();
            aSchemeColor6.Val = A.SchemeColorValues.PhColor;

            A.LuminanceModulation aLuminanceModulation5 = new A.LuminanceModulation();
            aLuminanceModulation5.Val = 99000;

            aSchemeColor6.Append(aLuminanceModulation5);

            A.SaturationModulation aSaturationModulation5 = new A.SaturationModulation();
            aSaturationModulation5.Val = 120000;

            aSchemeColor6.Append(aSaturationModulation5);

            A.Shade aShade1 = new A.Shade();
            aShade1.Val = 78000;

            aSchemeColor6.Append(aShade1);

            aGradientStop5.Append(aSchemeColor6);

            aGradientStopList1.Append(aGradientStop5);

            aGradientFill1.Append(aGradientStopList1);

            A.LinearGradientFill aLinearGradientFill1 = new A.LinearGradientFill();
            aLinearGradientFill1.Angle = 5400000;
            aLinearGradientFill1.Scaled = false;

            aGradientFill1.Append(aLinearGradientFill1);

            aFillStyleList.Append(aGradientFill1);

            aFormatScheme.Append(aFillStyleList);

            A.LineStyleList aLineStyleList = new A.LineStyleList();

            A.Outline aOutline = new A.Outline();
            aOutline.Width = 6350;
            aOutline.CapType = A.LineCapValues.Flat;
            aOutline.CompoundLineType = A.CompoundLineValues.Single;
            aOutline.Alignment = A.PenAlignmentValues.Center;

            A.SolidFill aSolidFill1 = new A.SolidFill();

            A.SchemeColor aSchemeColor7 = new A.SchemeColor();
            aSchemeColor7.Val = A.SchemeColorValues.PhColor;

            aSolidFill1.Append(aSchemeColor7);

            aOutline.Append(aSolidFill1);

            A.PresetDash aPresetDash = new A.PresetDash();
            aPresetDash.Val = A.PresetLineDashValues.Solid;

            aOutline.Append(aPresetDash);

            A.Miter aMiter = new A.Miter();
            aMiter.Limit = 800000;

            aOutline.Append(aMiter);

            aLineStyleList.Append(aOutline);

            A.Outline aOutline1 = new A.Outline();
            aOutline1.Width = 12700;
            aOutline1.CapType = A.LineCapValues.Flat;
            aOutline1.CompoundLineType = A.CompoundLineValues.Single;
            aOutline1.Alignment = A.PenAlignmentValues.Center;

            A.SolidFill aSolidFill2 = new A.SolidFill();

            A.SchemeColor aSchemeColor8 = new A.SchemeColor();
            aSchemeColor8.Val = A.SchemeColorValues.PhColor;

            aSolidFill2.Append(aSchemeColor8);

            aOutline1.Append(aSolidFill2);

            A.PresetDash aPresetDash1 = new A.PresetDash();
            aPresetDash1.Val = A.PresetLineDashValues.Solid;

            aOutline1.Append(aPresetDash1);

            A.Miter aMiter1 = new A.Miter();
            aMiter1.Limit = 800000;

            aOutline1.Append(aMiter1);

            aLineStyleList.Append(aOutline1);

            A.Outline aOutline2 = new A.Outline();
            aOutline2.Width = 19050;
            aOutline2.CapType = A.LineCapValues.Flat;
            aOutline2.CompoundLineType = A.CompoundLineValues.Single;
            aOutline2.Alignment = A.PenAlignmentValues.Center;

            A.SolidFill aSolidFill3 = new A.SolidFill();

            A.SchemeColor aSchemeColor9 = new A.SchemeColor();
            aSchemeColor9.Val = A.SchemeColorValues.PhColor;

            aSolidFill3.Append(aSchemeColor9);

            aOutline2.Append(aSolidFill3);

            A.PresetDash aPresetDash2 = new A.PresetDash();
            aPresetDash2.Val = A.PresetLineDashValues.Solid;

            aOutline2.Append(aPresetDash2);

            A.Miter aMiter2 = new A.Miter();
            aMiter2.Limit = 800000;

            aOutline2.Append(aMiter2);

            aLineStyleList.Append(aOutline2);

            aFormatScheme.Append(aLineStyleList);

            A.EffectStyleList aEffectStyleList = new A.EffectStyleList();

            A.EffectStyle aEffectStyle = new A.EffectStyle();

            A.EffectList aEffectList = new A.EffectList();

            aEffectStyle.Append(aEffectList);

            aEffectStyleList.Append(aEffectStyle);

            A.EffectStyle aEffectStyle1 = new A.EffectStyle();

            A.EffectList aEffectList1 = new A.EffectList();

            aEffectStyle1.Append(aEffectList1);

            aEffectStyleList.Append(aEffectStyle1);

            A.EffectStyle aEffectStyle2 = new A.EffectStyle();

            A.EffectList aEffectList2 = new A.EffectList();

            A.OuterShadow aOuterShadow = new A.OuterShadow();
            aOuterShadow.BlurRadius = 57150;
            aOuterShadow.Distance = 19050;
            aOuterShadow.Direction = 5400000;
            aOuterShadow.RotateWithShape = false;
            aOuterShadow.Alignment = A.RectangleAlignmentValues.Center;

            A.RgbColorModelHex aRgbColorModelHex10 = new A.RgbColorModelHex();
            aRgbColorModelHex10.Val = "000000";

            A.Alpha aAlpha = new A.Alpha();
            aAlpha.Val = 63000;

            aRgbColorModelHex10.Append(aAlpha);

            aOuterShadow.Append(aRgbColorModelHex10);

            aEffectList2.Append(aOuterShadow);

            aEffectStyle2.Append(aEffectList2);

            aEffectStyleList.Append(aEffectStyle2);

            aFormatScheme.Append(aEffectStyleList);

            A.BackgroundFillStyleList aBackgroundFillStyleList = new A.BackgroundFillStyleList();

            A.SolidFill aSolidFill4 = new A.SolidFill();

            A.SchemeColor aSchemeColor10 = new A.SchemeColor();
            aSchemeColor10.Val = A.SchemeColorValues.PhColor;

            aSolidFill4.Append(aSchemeColor10);

            aBackgroundFillStyleList.Append(aSolidFill4);

            A.SolidFill aSolidFill5 = new A.SolidFill();

            A.SchemeColor aSchemeColor11 = new A.SchemeColor();
            aSchemeColor11.Val = A.SchemeColorValues.PhColor;

            A.Tint aTint4 = new A.Tint();
            aTint4.Val = 95000;

            aSchemeColor11.Append(aTint4);

            A.SaturationModulation aSaturationModulation6 = new A.SaturationModulation();
            aSaturationModulation6.Val = 170000;

            aSchemeColor11.Append(aSaturationModulation6);

            aSolidFill5.Append(aSchemeColor11);

            aBackgroundFillStyleList.Append(aSolidFill5);

            A.GradientFill aGradientFill2 = new A.GradientFill();
            aGradientFill2.RotateWithShape = true;

            A.GradientStopList aGradientStopList2 = new A.GradientStopList();

            A.GradientStop aGradientStop6 = new A.GradientStop();
            aGradientStop6.Position = 0;

            A.SchemeColor aSchemeColor12 = new A.SchemeColor();
            aSchemeColor12.Val = A.SchemeColorValues.PhColor;

            A.Tint aTint5 = new A.Tint();
            aTint5.Val = 93000;

            aSchemeColor12.Append(aTint5);

            A.SaturationModulation aSaturationModulation7 = new A.SaturationModulation();
            aSaturationModulation7.Val = 150000;

            aSchemeColor12.Append(aSaturationModulation7);

            A.Shade aShade2 = new A.Shade();
            aShade2.Val = 98000;

            aSchemeColor12.Append(aShade2);

            A.LuminanceModulation aLuminanceModulation6 = new A.LuminanceModulation();
            aLuminanceModulation6.Val = 102000;

            aSchemeColor12.Append(aLuminanceModulation6);

            aGradientStop6.Append(aSchemeColor12);

            aGradientStopList2.Append(aGradientStop6);

            A.GradientStop aGradientStop7 = new A.GradientStop();
            aGradientStop7.Position = 50000;

            A.SchemeColor aSchemeColor13 = new A.SchemeColor();
            aSchemeColor13.Val = A.SchemeColorValues.PhColor;

            A.Tint aTint6 = new A.Tint();
            aTint6.Val = 98000;

            aSchemeColor13.Append(aTint6);

            A.SaturationModulation aSaturationModulation8 = new A.SaturationModulation();
            aSaturationModulation8.Val = 130000;

            aSchemeColor13.Append(aSaturationModulation8);

            A.Shade aShade3 = new A.Shade();
            aShade3.Val = 90000;

            aSchemeColor13.Append(aShade3);

            A.LuminanceModulation aLuminanceModulation7 = new A.LuminanceModulation();
            aLuminanceModulation7.Val = 103000;

            aSchemeColor13.Append(aLuminanceModulation7);

            aGradientStop7.Append(aSchemeColor13);

            aGradientStopList2.Append(aGradientStop7);

            A.GradientStop aGradientStop8 = new A.GradientStop();
            aGradientStop8.Position = 100000;

            A.SchemeColor aSchemeColor14 = new A.SchemeColor();
            aSchemeColor14.Val = A.SchemeColorValues.PhColor;

            A.Shade aShade4 = new A.Shade();
            aShade4.Val = 63000;

            aSchemeColor14.Append(aShade4);

            A.SaturationModulation aSaturationModulation9 = new A.SaturationModulation();
            aSaturationModulation9.Val = 120000;

            aSchemeColor14.Append(aSaturationModulation9);

            aGradientStop8.Append(aSchemeColor14);

            aGradientStopList2.Append(aGradientStop8);

            aGradientFill2.Append(aGradientStopList2);

            A.LinearGradientFill aLinearGradientFill2 = new A.LinearGradientFill();
            aLinearGradientFill2.Angle = 5400000;
            aLinearGradientFill2.Scaled = false;

            aGradientFill2.Append(aLinearGradientFill2);

            aBackgroundFillStyleList.Append(aGradientFill2);

            aFormatScheme.Append(aBackgroundFillStyleList);

            aThemeElements.Append(aFormatScheme);

            aTheme.Append(aThemeElements);

            A.ObjectDefaults aObjectDefaults = new A.ObjectDefaults();

            aTheme.Append(aObjectDefaults);

            A.ExtraColorSchemeList aExtraColorSchemeList = new A.ExtraColorSchemeList();

            aTheme.Append(aExtraColorSchemeList);

            A.OfficeStyleSheetExtensionList aOfficeStyleSheetExtensionList = new A.OfficeStyleSheetExtensionList();

            A.OfficeStyleSheetExtension aOfficeStyleSheetExtension = new A.OfficeStyleSheetExtension();
            aOfficeStyleSheetExtension.Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}";

            THM15.ThemeFamily thm15ThemeFamily = new THM15.ThemeFamily();

            thm15ThemeFamily.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

            thm15ThemeFamily.Name = "Office Theme";
            thm15ThemeFamily.Id = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}";
            thm15ThemeFamily.Vid = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}";

            aOfficeStyleSheetExtension.Append(thm15ThemeFamily);

            aOfficeStyleSheetExtensionList.Append(aOfficeStyleSheetExtension);

            aTheme.Append(aOfficeStyleSheetExtensionList);

            part.Theme = aTheme;
        }

        public void GenerateWorksheetPart(ref WorksheetPart part)
        {
            MarkupCompatibilityAttributes markupCompatibilityAttributes2 = new MarkupCompatibilityAttributes();
            markupCompatibilityAttributes2.Ignorable = "x14ac xr xr2 xr3";

            Worksheet worksheet = new Worksheet();

            worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            worksheet.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            worksheet.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
            worksheet.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");

            worksheet.MCAttributes = markupCompatibilityAttributes2;


            // Specify Sheet Dimensions
            // Create a 
            SheetDimension sheetDimension = new SheetDimension();
            sheetDimension.Reference = "A1:DL7";
            worksheet.Append(sheetDimension);


            SheetFormatProperties sheetFormatProperties = new SheetFormatProperties();
            sheetFormatProperties.DefaultRowHeight = 15D;
            sheetFormatProperties.DyDescent = 0.25D;

            worksheet.Append(sheetFormatProperties);

            /* *** Block the Creation of Columns***
            Columns columns = new Columns();

            Column column = new Column();
            column.Min = 30u;
            column.Max = 5u;
            column.Width = 9.7109375D;
            column.BestFit = true;
            column.CustomWidth = true;

            columns.Append(column);

            worksheet.Append(columns);
            */

            SheetData sheetData = new SheetData();

            ListValue<StringValue> listValueSv1 = new ListValue<StringValue>();
            listValueSv1.InnerText = "1:55";

            /*Row row = new Row();
            row.RowIndex = 1u;
            row.DyDescent = 0.25D;
            row.Spans = listValueSv1;*/



            /*            List<OpenXmlAttribute> oxa;
                        OpenXmlWriter oxw;

                        oxw = OpenXmlWriter.Create(part);


                        oxa = new List<OpenXmlAttribute>();
                        oxa.Add(new OpenXmlAttribute("r", null, "1"));
                        oxw.WriteStartElement(new Row(), oxa);


                        oxa = new List<OpenXmlAttribute>();
                        oxa.Add(new OpenXmlAttribute("t", null, "str"));
                        oxw.WriteStartElement(new Cell(), oxa);
                        oxw.WriteElement(new CellValue("Foo"));

                        oxw.WriteEndElement();
                        oxw.WriteEndElement();

                        oxw.Close();*/






            for (int y = 1; y <= 5; y++)
            {


                Row row = new Row();
                row.RowIndex = (uint)y;
                row.DyDescent = 0.25D;
                row.Spans = listValueSv1;

                if (y % 10 == 0)
                {
                    Console.WriteLine("Rows Written: " + y);
                }

                for (int x = 1; x <= 5; x++)
                {


                    Cell cell = new Cell();
                    cell.CellReference = $"{x}";
                    cell.StyleIndex = 1u;
                    cell.DataType = CellValues.String;
                    CellValue cellValue = new CellValue($"Row {y}");
                    cell.Append(cellValue);
                    row.Append(cell);

                }


                sheetData.Append(row);
            }






            /* ** Block original Cell Creation **
            Cell cell = new Cell();
            cell.CellReference = "A1";
            cell.StyleIndex = 1u;
            cell.DataType = CellValues.SharedString;

            CellValue cellValue = new CellValue("0");

            cell.Append(cellValue);

            row.Append(cell);

            Cell cell1 = new Cell();
            cell1.CellReference = "B1";
            cell1.StyleIndex = 1u;
            cell1.DataType = CellValues.SharedString;

            CellValue cellValue1 = new CellValue("1");

            cell1.Append(cellValue1);

            row.Append(cell1);

            Cell cell2 = new Cell();
            cell2.CellReference = "C1";
            cell2.StyleIndex = 1u;
            cell2.DataType = CellValues.SharedString;

            CellValue cellValue2 = new CellValue("2");

            cell2.Append(cellValue2);

            row.Append(cell2);

            Cell cell3 = new Cell();
            cell3.CellReference = "D1";
            cell3.StyleIndex = 1u;
            cell3.DataType = CellValues.SharedString;

            CellValue cellValue3 = new CellValue("3");

            cell3.Append(cellValue3);

            row.Append(cell3);

            Cell cell4 = new Cell();
            cell4.CellReference = "E1";
            cell4.StyleIndex = 1u;
            cell4.DataType = CellValues.SharedString;

            CellValue cellValue4 = new CellValue("4");

            cell4.Append(cellValue4);

            row.Append(cell4);


            ListValue<StringValue> listValueSv2 = new ListValue<StringValue>();
            listValueSv2.InnerText = "1:5";

            Row row1 = new Row();
            row1.RowIndex = 2u;
            row1.DyDescent = 0.25D;
            row1.Spans = listValueSv2;

            Cell cell5 = new Cell();
            cell5.CellReference = "A2";
            cell5.StyleIndex = 2u;

            CellValue cellValue5 = new CellValue("1");

            cell5.Append(cellValue5);

            row1.Append(cell5);

            Cell cell6 = new Cell();
            cell6.CellReference = "B2";
            cell6.StyleIndex = 3u;
            cell6.DataType = CellValues.SharedString;

            CellValue cellValue6 = new CellValue("5");

            cell6.Append(cellValue6);

            row1.Append(cell6);

            Cell cell7 = new Cell();
            cell7.CellReference = "C2";
            cell7.StyleIndex = 6u;
            cell7.DataType = CellValues.SharedString;

            CellValue cellValue7 = new CellValue("11");

            cell7.Append(cellValue7);

            row1.Append(cell7);

            Cell cell8 = new Cell();
            cell8.CellReference = "D2";
            cell8.StyleIndex = 4u;

            CellValue cellValue8 = new CellValue("6000");

            cell8.Append(cellValue8);

            row1.Append(cell8);

            Cell cell9 = new Cell();
            cell9.CellReference = "E2";
            cell9.StyleIndex = 5u;

            CellValue cellValue9 = new CellValue("44208");

            cell9.Append(cellValue9);

            row1.Append(cell9);

            sheetData.Append(row1);

            ListValue<StringValue> listValueSv3 = new ListValue<StringValue>();
            listValueSv3.InnerText = "1:5";

            Row row2 = new Row();
            row2.RowIndex = 3u;
            row2.DyDescent = 0.25D;
            row2.Spans = listValueSv3;

            Cell cell10 = new Cell();
            cell10.CellReference = "A3";
            cell10.StyleIndex = 2u;

            CellValue cellValue10 = new CellValue("2");

            cell10.Append(cellValue10);

            row2.Append(cell10);

            Cell cell11 = new Cell();
            cell11.CellReference = "B3";
            cell11.StyleIndex = 3u;
            cell11.DataType = CellValues.SharedString;

            CellValue cellValue11 = new CellValue("6");

            cell11.Append(cellValue11);

            row2.Append(cell11);

            Cell cell12 = new Cell();
            cell12.CellReference = "C3";
            cell12.StyleIndex = 6u;
            cell12.DataType = CellValues.SharedString;

            CellValue cellValue12 = new CellValue("11");

            cell12.Append(cellValue12);

            row2.Append(cell12);

            Cell cell13 = new Cell();
            cell13.CellReference = "D3";
            cell13.StyleIndex = 4u;

            CellValue cellValue13 = new CellValue("7000");

            cell13.Append(cellValue13);

            row2.Append(cell13);

            Cell cell14 = new Cell();
            cell14.CellReference = "E3";
            cell14.StyleIndex = 5u;

            CellValue cellValue14 = new CellValue("44209");

            cell14.Append(cellValue14);

            row2.Append(cell14);

            sheetData.Append(row2);

            ListValue<StringValue> listValueSv4 = new ListValue<StringValue>();
            listValueSv4.InnerText = "1:5";

            Row row3 = new Row();
            row3.RowIndex = 4u;
            row3.DyDescent = 0.25D;
            row3.Spans = listValueSv4;

            Cell cell15 = new Cell();
            cell15.CellReference = "A4";
            cell15.StyleIndex = 2u;

            CellValue cellValue15 = new CellValue("3");

            cell15.Append(cellValue15);

            row3.Append(cell15);

            Cell cell16 = new Cell();
            cell16.CellReference = "B4";
            cell16.StyleIndex = 3u;
            cell16.DataType = CellValues.SharedString;

            CellValue cellValue16 = new CellValue("7");

            cell16.Append(cellValue16);

            row3.Append(cell16);

            Cell cell17 = new Cell();
            cell17.CellReference = "C4";
            cell17.StyleIndex = 6u;
            cell17.DataType = CellValues.SharedString;

            CellValue cellValue17 = new CellValue("11");

            cell17.Append(cellValue17);

            row3.Append(cell17);

            Cell cell18 = new Cell();
            cell18.CellReference = "D4";
            cell18.StyleIndex = 4u;

            CellValue cellValue18 = new CellValue("3000");

            cell18.Append(cellValue18);

            row3.Append(cell18);

            Cell cell19 = new Cell();
            cell19.CellReference = "E4";
            cell19.StyleIndex = 5u;

            CellValue cellValue19 = new CellValue("44210");

            cell19.Append(cellValue19);

            row3.Append(cell19);

            sheetData.Append(row3);

            ListValue<StringValue> listValueSv5 = new ListValue<StringValue>();
            listValueSv5.InnerText = "1:5";

            Row row4 = new Row();
            row4.RowIndex = 5u;
            row4.DyDescent = 0.25D;
            row4.Spans = listValueSv5;

            Cell cell20 = new Cell();
            cell20.CellReference = "A5";
            cell20.StyleIndex = 2u;

            CellValue cellValue20 = new CellValue("3");

            cell20.Append(cellValue20);

            row4.Append(cell20);

            Cell cell21 = new Cell();
            cell21.CellReference = "B5";
            cell21.StyleIndex = 3u;
            cell21.DataType = CellValues.SharedString;

            CellValue cellValue21 = new CellValue("8");

            cell21.Append(cellValue21);

            row4.Append(cell21);

            Cell cell22 = new Cell();
            cell22.CellReference = "C5";
            cell22.StyleIndex = 6u;
            cell22.DataType = CellValues.SharedString;

            CellValue cellValue22 = new CellValue("11");

            cell22.Append(cellValue22);

            row4.Append(cell22);

            Cell cell23 = new Cell();
            cell23.CellReference = "D5";
            cell23.StyleIndex = 4u;

            CellValue cellValue23 = new CellValue("4500");

            cell23.Append(cellValue23);

            row4.Append(cell23);

            Cell cell24 = new Cell();
            cell24.CellReference = "E5";
            cell24.StyleIndex = 5u;

            CellValue cellValue24 = new CellValue("44211");

            cell24.Append(cellValue24);

            row4.Append(cell24);

            sheetData.Append(row4);

            ListValue<StringValue> listValueSv6 = new ListValue<StringValue>();
            listValueSv6.InnerText = "1:5";

            Row row5 = new Row();
            row5.RowIndex = 6u;
            row5.DyDescent = 0.25D;
            row5.Spans = listValueSv6;

            Cell cell25 = new Cell();
            cell25.CellReference = "A6";
            cell25.StyleIndex = 2u;

            CellValue cellValue25 = new CellValue("4");

            cell25.Append(cellValue25);

            row5.Append(cell25);

            Cell cell26 = new Cell();
            cell26.CellReference = "B6";
            cell26.StyleIndex = 3u;
            cell26.DataType = CellValues.SharedString;

            CellValue cellValue26 = new CellValue("9");

            cell26.Append(cellValue26);

            row5.Append(cell26);

            Cell cell27 = new Cell();
            cell27.CellReference = "C6";
            cell27.StyleIndex = 6u;
            cell27.DataType = CellValues.SharedString;

            CellValue cellValue27 = new CellValue("11");

            cell27.Append(cellValue27);

            row5.Append(cell27);

            Cell cell28 = new Cell();
            cell28.CellReference = "D6";
            cell28.StyleIndex = 4u;

            CellValue cellValue28 = new CellValue("17");

            cell28.Append(cellValue28);

            row5.Append(cell28);

            Cell cell29 = new Cell();
            cell29.CellReference = "E6";
            cell29.StyleIndex = 5u;

            CellValue cellValue29 = new CellValue("44212");

            cell29.Append(cellValue29);

            row5.Append(cell29);

            sheetData.Append(row5);

            ListValue<StringValue> listValueSv7 = new ListValue<StringValue>();
            listValueSv7.InnerText = "1:5";

            Row row6 = new Row();
            row6.RowIndex = 7u;
            row6.DyDescent = 0.25D;
            row6.Spans = listValueSv7;

            Cell cell30 = new Cell();
            cell30.CellReference = "A7";
            cell30.StyleIndex = 2u;

            CellValue cellValue30 = new CellValue("5");

            cell30.Append(cellValue30);

            row6.Append(cell30);

            Cell cell31 = new Cell();
            cell31.CellReference = "B7";
            cell31.StyleIndex = 3u;
            cell31.DataType = CellValues.SharedString;

            CellValue cellValue31 = new CellValue("10");

            cell31.Append(cellValue31);

            row6.Append(cell31);

            Cell cell32 = new Cell();
            cell32.CellReference = "C7";
            cell32.StyleIndex = 6u;
            cell32.DataType = CellValues.SharedString;

            CellValue cellValue32 = new CellValue("11");

            cell32.Append(cellValue32);

            row6.Append(cell32);

            Cell cell33 = new Cell();
            cell33.CellReference = "D7";
            cell33.StyleIndex = 4u;

            CellValue cellValue33 = new CellValue("18");

            cell33.Append(cellValue33);

            row6.Append(cell33);

            Cell cell34 = new Cell();
            cell34.CellReference = "E7";
            cell34.StyleIndex = 5u;

            CellValue cellValue34 = new CellValue("44213");

            cell34.Append(cellValue34);

            row6.Append(cell34);

            sheetData.Append(row6);*/

            worksheet.Append(sheetData);

            PhoneticProperties phoneticProperties = new PhoneticProperties();
            phoneticProperties.FontId = 2u;
            phoneticProperties.Type = PhoneticValues.NoConversion;

            worksheet.Append(phoneticProperties);

            PageMargins pageMargins = new PageMargins();
            pageMargins.Left = 0.7D;
            pageMargins.Right = 0.7D;
            pageMargins.Top = 0.75D;
            pageMargins.Bottom = 0.75D;
            pageMargins.Header = 0.3D;
            pageMargins.Footer = 0.3D;

            worksheet.Append(pageMargins);

            PageSetup pageSetup = new PageSetup();
            pageSetup.Id = "rId1";
            pageSetup.Orientation = OrientationValues.Portrait;

            worksheet.Append(pageSetup);

            part.Worksheet = worksheet;
        }

        public void GenerateSpreadsheetPrinterSettingsPart(ref SpreadsheetPrinterSettingsPart part)
        {
            string base64 = "SABQADYAMABGAEQAMABEACAAKABIAFAAIABEAGUAcwBrAEoAZQB0ACAAMwA2ADMAMAAgAHMAZQByAGkAA" +
                "AAAAAEEAAbcACwIQ78BAgEAAQDqCm8IZAABAA8AWAICAAEAWAIDAAEATABlAHQAdABlAHIAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "QAAAAAAAAABAAAAAgAAAAgBAAD/////AAAAAAAAAAAAAAAAAAAAAERJTlUiALAF7AdAAGDew/YAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAEAAAAAAAAACQABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "QAAAAAAAAAAAAAAsAUAAFNNVEoAAAAAEACgBQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAABNU0RYAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=";

            Stream mem = new MemoryStream(Convert.FromBase64String(base64), false);
            try
            {
                part.FeedData(mem);
            }
            finally
            {
                mem.Dispose();
            }
        }

        public void GenerateSharedStringTablePart(ref SharedStringTablePart part)
        {
            SharedStringTable sharedStringTable = new SharedStringTable();
            sharedStringTable.Count = 55u;
            sharedStringTable.UniqueCount = 1u;

            //SharedStringItem sharedStringItem = new SharedStringItem();


            /*            for (int x = 1; x <= 55; x++)
                        {
                            SharedStringItem sharedStringItem = new SharedStringItem();
                            Text text = new Text("Column");
                            sharedStringItem.Append(text);
                            sharedStringTable.Append(sharedStringItem);

                        }*/


            /* Writes three columns successfully
            SharedStringItem sharedStringItem1 = new SharedStringItem();
            Text text1 = new Text("Column");
            sharedStringItem1.Append(text1);
            sharedStringTable.Append(sharedStringItem1);

            SharedStringItem sharedStringItem2 = new SharedStringItem();
            Text text2 = new Text("Column");
            sharedStringItem2.Append(text2);
            sharedStringTable.Append(sharedStringItem2);


            SharedStringItem sharedStringItem3 = new SharedStringItem();
            Text text3 = new Text("Column");
            sharedStringItem3.Append(text3);
            sharedStringTable.Append(sharedStringItem3);*/



            // Text text = new Text("ID");

            //sharedStringItem.Append(text);

            /*  sharedStringTable.Append(sharedStringItem);

              SharedStringItem sharedStringItem1 = new SharedStringItem();

              Text text1 = new Text("TEST");

              sharedStringItem1.Append(text1);

              sharedStringTable.Append(sharedStringItem1);

              SharedStringItem sharedStringItem2 = new SharedStringItem();

              Text text2 = new Text("NEW");

              sharedStringItem2.Append(text2);

              sharedStringTable.Append(sharedStringItem2);

              SharedStringItem sharedStringItem3 = new SharedStringItem();

              Text text3 = new Text("FOO");

              sharedStringItem3.Append(text3);

              sharedStringTable.Append(sharedStringItem3);

              SharedStringItem sharedStringItem4 = new SharedStringItem();

              Text text4 = new Text("BAR");

              sharedStringItem4.Append(text4);

              sharedStringTable.Append(sharedStringItem4);*/










            /*SharedStringItem sharedStringItem5 = new SharedStringItem();

            Text text5 = new Text("00011");

            sharedStringItem5.Append(text5);

            sharedStringTable.Append(sharedStringItem5);

            SharedStringItem sharedStringItem6 = new SharedStringItem();

            Text text6 = new Text("00012");

            sharedStringItem6.Append(text6);

            sharedStringTable.Append(sharedStringItem6);

            SharedStringItem sharedStringItem7 = new SharedStringItem();

            Text text7 = new Text("00013");

            sharedStringItem7.Append(text7);

            sharedStringTable.Append(sharedStringItem7);

            SharedStringItem sharedStringItem8 = new SharedStringItem();

            Text text8 = new Text("00014");

            sharedStringItem8.Append(text8);

            sharedStringTable.Append(sharedStringItem8);

            SharedStringItem sharedStringItem9 = new SharedStringItem();

            Text text9 = new Text("00015");

            sharedStringItem9.Append(text9);

            sharedStringTable.Append(sharedStringItem9);

            SharedStringItem sharedStringItem10 = new SharedStringItem();

            Text text10 = new Text("00016");

            sharedStringItem10.Append(text10);

            sharedStringTable.Append(sharedStringItem10);

            SharedStringItem sharedStringItem11 = new SharedStringItem();

            Text text11 = new Text("WWW");

            sharedStringItem11.Append(text11);

            sharedStringTable.Append(sharedStringItem11);*/

            part.SharedStringTable = sharedStringTable;
        }

    }
}