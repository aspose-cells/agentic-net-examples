using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfSaveDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Populate the first worksheet with sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Aspose.Cells PDF Save Demo");
        sheet.Cells["A2"].PutValue(DateTime.Now);
        sheet.Cells["B1"].PutValue(123);
        sheet.Cells["B2"].Formula = "=SUM(B1)";

        // Create PDF save options and configure desired settings
        PdfSaveOptions options = new PdfSaveOptions();
        options.Compliance = PdfCompliance.PdfA1b;               // Set PDF/A-1b compliance
        options.CalculateFormula = true;                        // Calculate formulas before saving
        options.OptimizationType = PdfOptimizationType.MinimumSize; // Optimize for minimum file size
        options.OnePagePerSheet = true;                         // Render each sheet on a single page
        options.AllColumnsInOnePagePerSheet = true;             // Fit all columns on one page
        options.CheckFontCompatibility = true;                  // Ensure font compatibility
        options.DefaultFont = "Arial";                          // Default font for Unicode characters

        // Save the workbook to PDF using the Save(string, SaveOptions) rule
        string pdfPath = "DemoOutput.pdf";
        workbook.Save(pdfPath, options);

        // Verify that the PDF file was created and has content
        if (File.Exists(pdfPath))
        {
            long fileSize = new FileInfo(pdfPath).Length;
            Console.WriteLine($"PDF saved successfully. File size: {fileSize} bytes.");
        }
        else
        {
            Console.WriteLine("Failed to create the PDF file.");
        }
    }
}