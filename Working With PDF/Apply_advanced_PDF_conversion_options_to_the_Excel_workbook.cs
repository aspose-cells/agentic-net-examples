using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering.PdfSecurity;
using System.Drawing;

class AdvancedPdfConversion
{
    static void Main()
    {
        // Create a new workbook and populate it with sample data
        Workbook workbook = new Workbook();
        Worksheet ws = workbook.Worksheets[0];
        ws.Cells["A1"].PutValue("Name");
        ws.Cells["B1"].PutValue("Age");
        ws.Cells["A2"].PutValue("John");
        ws.Cells["B2"].PutValue(30);
        ws.Cells["A3"].PutValue("Alice");
        ws.Cells["B3"].PutValue(25);

        // Initialize PDF save options and set advanced conversion properties
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // PDF compliance (PDF/A-1b)
        pdfOptions.Compliance = PdfCompliance.PdfA1b;

        // Embed standard Windows fonts
        pdfOptions.EmbedStandardWindowsFonts = true;

        // Calculate formulas before saving
        pdfOptions.CalculateFormula = true;

        // Use Flate compression for PDF streams
        pdfOptions.PdfCompression = PdfCompressionCore.Flate;

        // Export custom document properties in standard way
        pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

        // Preserve document structure for accessibility
        pdfOptions.ExportDocumentStructure = true;

        // Show document title in PDF viewer window
        pdfOptions.DisplayDocTitle = true;

        // Default font for Unicode characters
        pdfOptions.DefaultFont = "Arial";

        // Use workbook's default font when cell font is missing
        pdfOptions.CheckWorkbookDefaultFont = true;

        // Verify font compatibility for each character
        pdfOptions.CheckFontCompatibility = true;

        // Substitute fonts at character granularity only
        pdfOptions.IsFontSubstitutionCharGranularity = true;

        // Render each sheet on a single page
        pdfOptions.OnePagePerSheet = true;
        pdfOptions.AllColumnsInOnePagePerSheet = true;

        // Hide rendering errors
        pdfOptions.IgnoreError = true;

        // Output a blank page if nothing to print
        pdfOptions.OutputBlankPageWhenNothingToPrint = true;

        // Save only the first page (0‑based index) and limit to one page
        pdfOptions.PageIndex = 0;
        pdfOptions.PageCount = 1;

        // Ignore blank pages during printing
        pdfOptions.PrintingPageType = PrintingPageType.IgnoreBlank;

        // Use dotted gridlines
        pdfOptions.GridlineType = GridlineType.Dotted;

        // Keep text crossing cell boundaries
        pdfOptions.TextCrossType = TextCrossType.CrossKeep;

        // Set default edit language
        pdfOptions.DefaultEditLanguage = DefaultEditLanguage.English;

        // Render all visible sheets
        pdfOptions.SheetSet = SheetSet.Visible;

        // Do not clear data after saving
        pdfOptions.ClearData = false;

        // Temporary folder for cache files
        pdfOptions.CachedFileFolder = "C:\\Temp";

        // Validate merged cells and merge conditional formatting areas
        pdfOptions.ValidateMergedAreas = true;
        pdfOptions.MergeAreas = true;

        // Sort defined names and external names
        pdfOptions.SortNames = true;
        pdfOptions.SortExternalNames = true;

        // Refresh chart cache before rendering
        pdfOptions.RefreshChartCache = true;

        // Optional: set PDF security options
        PdfSecurityOptions security = new PdfSecurityOptions
        {
            OwnerPassword = "ownerPwd",
            UserPassword = "userPwd",
            PrintPermission = true,
            FullQualityPrintPermission = true
        };
        pdfOptions.SecurityOptions = security;

        // Optional: add a watermark to the PDF
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

        // Save the workbook as a PDF using the configured options
        workbook.Save("AdvancedOutput.pdf", pdfOptions);
    }
}