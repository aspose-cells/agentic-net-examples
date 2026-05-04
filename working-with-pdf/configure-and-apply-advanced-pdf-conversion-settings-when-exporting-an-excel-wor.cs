using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfAdvancedExport
{
    public class AdvancedPdfExport
    {
        public static void Run()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Report";

            // Populate cells with sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["C2"].PutValue(0.5);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["C3"].PutValue(0.3);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["C4"].PutValue(0.8);

            // Add custom document properties
            workbook.CustomDocumentProperties.Add("Author", "John Doe");
            workbook.CustomDocumentProperties.Add("Department", "Sales");

            // Create PDF save options and configure advanced settings
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // PDF/A compliance (PDF/A-1b)
                Compliance = PdfCompliance.PdfA1b,

                // Embed standard Windows fonts
                EmbedStandardWindowsFonts = true,

                // Export custom document properties to PDF
                CustomPropertiesExport = PdfCustomPropertiesExport.Standard,

                // Export document structure (tags) for accessibility
                ExportDocumentStructure = true,

                // Show document title in viewer title bar
                DisplayDocTitle = true,

                // Set font encoding to Identity (Unicode)
                FontEncoding = PdfFontEncoding.Identity,

                // Embed attachments (e.g., OLE objects) into PDF
                EmbedAttachments = true,

                // Default font to use when original font is missing
                DefaultFont = "Arial",

                // Use workbook's default font when characters are Unicode but font missing
                CheckWorkbookDefaultFont = true,

                // Check font compatibility for each character
                CheckFontCompatibility = true,

                // Substitute font at character granularity only
                IsFontSubstitutionCharGranularity = true,

                // Render each sheet on a single page
                OnePagePerSheet = true,

                // Fit all columns of a sheet onto one page
                AllColumnsInOnePagePerSheet = true,

                // Ignore rendering errors (e.g., unsupported shapes)
                IgnoreError = true,

                // Output a blank page if a sheet has nothing to print
                OutputBlankPageWhenNothingToPrint = true,

                // Save only the first page (example)
                PageIndex = 0,
                PageCount = 1,

                // Ignore blank pages during printing
                PrintingPageType = PrintingPageType.IgnoreBlank,

                // Set gridline appearance
                GridlineType = GridlineType.Dotted,
                GridlineColor = Color.LightGray,

                // Text handling when it exceeds cell width
                TextCrossType = TextCrossType.CrossKeep,

                // Default edit language
                DefaultEditLanguage = DefaultEditLanguage.English,

                // Render only visible sheets
                SheetSet = SheetSet.Visible,

                // Clear workbook data after saving (optional)
                ClearData = false,

                // Temporary folder for cache files
                CachedFileFolder = @"C:\Temp\AsposePdfCache",

                // Validate merged areas before saving
                ValidateMergedAreas = true,

                // Merge conditional formatting and validation areas
                MergeAreas = true,

                // Sort defined names and external names
                SortNames = true,
                SortExternalNames = true,

                // Refresh chart cache before rendering
                RefreshChartCache = true
            };

            // Add a watermark to the PDF
            RenderingFont watermarkFont = new RenderingFont("Calibri", 68)
            {
                Italic = true,
                Bold = true,
                Color = Color.Blue
            };
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
            {
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                Rotation = 30,
                Opacity = 0.6f,
                ScaleToPagePercent = 50
            };
            pdfOptions.Watermark = watermark;

            // Save the workbook as PDF with the configured options
            workbook.Save("AdvancedReport.pdf", pdfOptions);

            Console.WriteLine("Advanced PDF export completed successfully.");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            AdvancedPdfExport.Run();
        }
    }
}