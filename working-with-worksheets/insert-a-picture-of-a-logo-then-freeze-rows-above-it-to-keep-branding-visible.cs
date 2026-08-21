// Title: Insert a logo image and freeze top rows in an Excel worksheet using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, conditionally adds a PNG logo to cell B2, freezes the first three rows to keep the branding visible, and saves the file as BrandingWorkbook.xlsx.
// Keywords: Aspose.Cells insert image C# | freeze panes Aspose.Cells | add logo to Excel workbook | Excel header branding Aspose | C# picture insertion Excel | Aspose.Cells FreezePanes example | Excel workbook branding code
// Common Searches: Aspose.Cells C# add logo to Excel | How to freeze rows after inserting a picture with Aspose.Cells | Insert image and freeze top rows in Excel using Aspose.Cells | C# code to keep Excel header visible with Aspose.Cells | Aspose.Cells picture insertion and FreezePanes
// Developer Intent: Place a logo in the worksheet header and keep it visible while scrolling by freezing the rows above it.
// Use Cases: Design a branded report template where the company logo stays fixed during data navigation. | Generate invoices with a persistent logo header for consistent corporate identity. | Automate dashboards that require a static branding banner while users scroll through large datasets.
// AI Prompts: Generate C# code with Aspose.Cells that inserts a PNG logo at B2 and freezes the first three rows. | Show an Aspose.Cells example that checks for a logo file, adds it to a worksheet, and applies FreezePanes to keep the header visible. | Explain how to modify the freeze point to keep additional rows visible after adding a picture with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Creates a new workbook, conditionally adds a PNG logo to cell B2, freezes the first three rows to keep the branding visible, and saves the file as BrandingWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the logo image
            string logoFile = "logo.png";

            // Insert the logo picture at cell B2 (row index 1, column index 1) if the file exists
            if (File.Exists(logoFile))
            {
                worksheet.Pictures.Add(1, 1, logoFile);
            }
            else
            {
                Console.WriteLine($"Warning: Logo file '{logoFile}' not found. Skipping picture insertion.");
            }

            // Freeze the first three rows (rows 0‑2). Freeze point is set at row index 3.
            worksheet.FreezePanes(3, 0, 3, 0);

            // Save the workbook to a file
            string outputFile = "BrandingWorkbook.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
