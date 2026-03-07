using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Rendering;

class ConvertExcelToPdfWithAttributes
{
    static void Main()
    {
        // Create a sample workbook with some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello Aspose");
        sheet.Cells["A2"].PutValue("PDF conversion with attributes");

        // Save the workbook to a temporary Excel file (source for conversion)
        string sourcePath = "sample.xlsx";
        workbook.Save(sourcePath);

        // Load options for the source file (optional, here we specify XLSX format)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

        // Configure PDF save options with desired conversion attributes
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set PDF/A-1b compliance
            Compliance = PdfCompliance.PdfA1b,

            // Optimize for minimum file size
            OptimizationType = PdfOptimizationType.MinimumSize,

            // Export custom document properties as standard entries
            CustomPropertiesExport = PdfCustomPropertiesExport.Standard,

            // Render each worksheet on a single page
            OnePagePerSheet = true,

            // Use workbook's default font when a cell's font is missing
            CheckWorkbookDefaultFont = true,

            // Fallback font if needed
            DefaultFont = "Arial",

            // Embed standard Windows fonts in the PDF
            EmbedStandardWindowsFonts = true
        };

        // Destination PDF file path
        string destPath = "output.pdf";

        // Perform the conversion using ConversionUtility with the specified options
        ConversionUtility.Convert(sourcePath, loadOptions, destPath, pdfOptions);

        Console.WriteLine("Excel file has been successfully converted to PDF with the specified attributes.");
    }
}