using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class CustomPdfMarginDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data to visualize the margins
        sheet.Cells["A1"].PutValue("Custom PDF Margin Demo");
        sheet.Cells["A2"].PutValue("Left, Right, Top, Bottom margins are set via PageSetup.");

        // Set custom page margins (values are in centimeters)
        // Example: 2 cm left, 2 cm right, 3 cm top, 1.5 cm bottom
        sheet.PageSetup.LeftMargin = 2.0;
        sheet.PageSetup.RightMargin = 2.0;
        sheet.PageSetup.TopMargin = 3.0;
        sheet.PageSetup.BottomMargin = 1.5;

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // (Optional) Set additional options, e.g., optimize for minimum size
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as PDF with the custom margin settings
        string outputPath = "CustomMarginsDemo.pdf";
        workbook.Save(outputPath, pdfOptions);

        Console.WriteLine($"PDF saved to '{outputPath}' with custom page margins.");
    }
}