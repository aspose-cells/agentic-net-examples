using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // 600 points = 600 / 72 inches (1 point = 1/72 inch)
        double sizeInInches = 600.0 / 72.0; // ≈ 8.3333 inches

        // Apply a custom paper size (width and height) using the PageSetup.CustomPaperSize method
        worksheet.PageSetup.CustomPaperSize(sizeInInches, sizeInInches);
        // Set the PaperSize to Custom so the custom dimensions are used
        worksheet.PageSetup.PaperSize = PaperSizeType.Custom;

        // Render the worksheet to PDF by saving the workbook in PDF format
        workbook.Save("CustomPaperSize.pdf", SaveFormat.Pdf);

        // Revert the worksheet's paper size back to standard A4
        worksheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

        // Save the workbook after reverting (optional, e.g., as an Excel file)
        workbook.Save("RevertedToA4.xlsx");
    }
}