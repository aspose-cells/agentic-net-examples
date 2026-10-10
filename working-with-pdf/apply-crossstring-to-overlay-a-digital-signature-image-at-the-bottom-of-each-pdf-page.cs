// Title: Add a PNG digital signature to the bottom of each worksheet and export to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that inserts a PNG signature at the last used row of every worksheet, sets its dimensions, and saves the workbook as a PDF with Aspose.Cells. | Create a reusable C# method that adds a scaled signature picture to each sheet's footer area before converting the workbook to PDF using Aspose.Cells. | Adjust the example to position the signature image in the PDF page footer by configuring picture placement and worksheet margins with Aspose.Cells.
// Common Searches: how to add a digital signature image to each page when converting Excel to PDF with Aspose.Cells C# | C# Aspose.Cells place picture at bottom of worksheet before PDF export | scale PNG signature to points and embed in PDF using Aspose.Cells .NET | move and size picture placement Aspose.Cells for PDF generation
// Tags: add signature image Aspose.Cells PDF export | insert picture bottom row worksheet C# | scale PNG to points Aspose.Cells | picture placement MoveAndSize Aspose.Cells | automate digital signature overlay PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample checks for the Excel workbook and PNG signature files, loads the workbook, calculates the signature size, inserts the image at the first column of the last used row on each worksheet with MoveAndSize placement, and then saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        try
        {
            // Paths
            string excelPath = "input.xlsx";
            string signatureImagePath = "signature.png";
            string outputPdfPath = "output_signed.pdf";

            // Verify input files exist
            if (!File.Exists(excelPath))
                throw new FileNotFoundException($"Excel file not found: {excelPath}");
            if (!File.Exists(signatureImagePath))
                throw new FileNotFoundException($"Signature image not found: {signatureImagePath}");

            // Load workbook
            Workbook workbook = new Workbook(excelPath);

            // Define signature size (points; 1 point = 1/72 inch)
            const double sigWidthPoints = 150;
            const double sigHeightPoints = 50;

            // Convert points to pixels (approx. 96 DPI)
            int sigWidthPixels = (int)(sigWidthPoints * 96 / 72);
            int sigHeightPixels = (int)(sigHeightPoints * 96 / 72);

            // Add signature image to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the bottom row of the used range
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                int bottomRow = usedRange != null
                    ? usedRange.FirstRow + usedRange.RowCount
                    : 0; // if sheet is empty, start at first row

                // Add picture at the first column of the bottom row
                int picIndex = sheet.Pictures.Add(bottomRow, 0, signatureImagePath);
                Picture pic = sheet.Pictures[picIndex];

                // Set size (pixels)
                pic.Width = sigWidthPixels;
                pic.Height = sigHeightPixels;

                // Ensure the picture moves and resizes with cells
                pic.Placement = PlacementType.MoveAndSize;
            }

            // Save as PDF
            workbook.Save(outputPdfPath, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
