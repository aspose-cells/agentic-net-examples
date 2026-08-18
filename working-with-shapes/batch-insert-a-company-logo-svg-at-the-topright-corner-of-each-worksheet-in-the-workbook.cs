// Title: C# – Batch Insert a Company Logo SVG into the Top‑Right Corner of Every Worksheet with Aspose.Cells
// Description: Loads a workbook, reads a company_logo.svg file, and uses Aspose.Cells' AddSvg method to place the SVG at the first row and a right‑hand column on each worksheet. The shape auto‑sizes and the workbook is saved as output.xlsx.
// Keywords: Aspose.Cells | C# | AddSvg | insert SVG | company logo | top right corner | batch insert | multiple worksheets | Excel automation | shape placement | Excel workbook | SVG shape | code example | Aspose.Cells for .NET | Excel branding
// Common Searches: C# Aspose.Cells add SVG to all sheets | How to place a logo in the top‑right of each worksheet using Aspose.Cells | Batch insert picture in Excel workbook Aspose.Cells .NET | AddSvg method example multiple worksheets | Insert company logo SVG Aspose.Cells C#
// Developer Intent: Add the same SVG logo to the top‑right corner of every worksheet in an Excel workbook.
// Use Cases: Apply consistent branding to all sheets of a generated report. | Insert a seal or trademark SVG across a multi‑sheet financial model. | Create a reusable template that automatically adds a header logo. | Prepare multi‑sheet invoices with the company logo pre‑placed. | Automate compliance watermark insertion on every worksheet.
// AI Prompts: Write C# code that loads an SVG file and uses Aspose.Cells to add it as a shape to the top‑right cell of each worksheet, including file‑not‑found handling. | Show how to adjust column index and offset values in AddSvg so the logo aligns with the right margin on sheets of different widths. | Provide an example that saves the workbook in both .xlsx and .xlsb formats after inserting the SVG logo on all worksheets. | Generate a reusable method that accepts a workbook and SVG path, then batches the logo insertion across all worksheets.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, reads a company_logo.svg file, and uses Aspose.Cells' AddSvg method to place the SVG at the first row and a right‑hand column on each worksheet. The shape auto‑sizes and the workbook is saved as output.xlsx.
class InsertSvgLogo
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one if needed)
            Workbook workbook = new Workbook();

            // Path to the SVG logo file
            string svgPath = "company_logo.svg";

            // Verify that the SVG file exists before attempting to read it
            if (!File.Exists(svgPath))
            {
                Console.WriteLine($"SVG file not found: {svgPath}");
                return;
            }

            // Load the SVG logo into a byte array
            byte[] svgData = File.ReadAllBytes(svgPath);

            // Insert the SVG into the top‑right corner of every worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                ShapeCollection shapes = sheet.Shapes;

                // Position parameters
                int topRow = 0;          // first row
                int top = 0;             // vertical offset
                int leftColumn = 10;     // column index near the right edge
                int left = 0;            // horizontal offset
                int height = -1;         // auto‑size height
                int width = -1;          // auto‑size width

                // Add the SVG picture (compatibleImageData is null for newer Excel versions)
                shapes.AddSvg(topRow, top, leftColumn, left, height, width, svgData, null);
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
