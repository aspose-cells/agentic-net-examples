// Title: Set a 600‑point custom page size for a worksheet, export to PDF, then revert to A4 with Aspose.Cells for .NET (C#)
// Common Searches: c# aspocells set worksheet custom page size to 600 points | how to export a worksheet as PDF with a 600x600 point page using Aspose.Cells | reset worksheet page setup to A4 after using a custom size in Aspose.Cells C# | Aspose.Cells C# generate PDF with specific point dimensions for page layout | save workbook to PDF after changing page size then revert to default size Aspose.Cells
// Tags: custom page dimensions points Aspose.Cells | worksheet page setup custom size C# | export worksheet to PDF Aspose.Cells | reset page setup to A4 Aspose.Cells | save workbook as PDF with custom size | paper size type custom Aspose.Cells

using System;
using Aspose.Cells;
using System.Drawing;

// The example creates a workbook, configures the first worksheet's PageSetup to a custom 600 × 600‑point size, saves it as a PDF, then switches the PaperSize back to A4 and saves a second PDF.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Apply custom paper size (600 points x 600 points) -----
            // Set the paper size to Custom
            sheet.PageSetup.PaperSize = PaperSizeType.Custom;
            // Define the custom size in points (1 point = 1/72 inch)
            sheet.PageSetup.CustomPaperSize(600.0, 600.0);

            // Optionally add some content to visualize the page size
            sheet.Cells["A1"].PutValue("Custom 600‑point paper size");

            // Render the worksheet to PDF with the custom paper size
            workbook.Save("CustomPaperSize.pdf", SaveFormat.Pdf);

            // ----- Revert to A4 paper size -----
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

            // Render the worksheet to PDF with A4 size
            workbook.Save("A4PaperSize.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
