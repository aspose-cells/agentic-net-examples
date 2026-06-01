using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Sample Data");
            worksheet.Cells["A2"].PutValue(123);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Compliance and metadata
                Compliance = PdfCompliance.PdfA1b,
                Producer = "Aspose.Cells Test",
                CreatedTime = DateTime.Now,

                // Optimization and compression
                OptimizationType = PdfOptimizationType.MinimumSize,
                PdfCompression = PdfCompressionCore.Flate,

                // Font handling
                CheckFontCompatibility = true,
                CheckWorkbookDefaultFont = true,
                DefaultFont = "Arial",
                EmbedStandardWindowsFonts = true,
                EmbedAttachments = true,

                // Layout
                OnePagePerSheet = true,
                AllColumnsInOnePagePerSheet = true,
                GridlineType = GridlineType.Dotted,
                TextCrossType = TextCrossType.CrossKeep,

                // Page selection
                PageIndex = 0,
                PageCount = 1,
                PrintingPageType = PrintingPageType.IgnoreBlank,

                // Miscellaneous
                ExportDocumentStructure = true,
                CalculateFormula = true
            };

            // Save the workbook as a PDF file using the configured options
            string pdfPath = "output.pdf";

            // Ensure the target directory exists
            string pdfDir = Path.GetDirectoryName(Path.GetFullPath(pdfPath));
            if (!Directory.Exists(pdfDir))
                Directory.CreateDirectory(pdfDir);

            workbook.Save(pdfPath, pdfOptions);

            // Verify that the PDF file was created
            bool fileCreated = File.Exists(pdfPath);
            Console.WriteLine($"PDF creation verification: {(fileCreated ? "Passed" : "Failed")}");

            // Additional simple verification: file size should be greater than zero
            if (fileCreated)
            {
                long fileSize = new FileInfo(pdfPath).Length;
                Console.WriteLine($"PDF size verification: {(fileSize > 0 ? "Passed" : "Failed")} (Size: {fileSize} bytes)");
            }
        }
        catch (Exception ex)
        {
            // Output any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}