// Title: Batch insert an SVG logo into the top‑right corner of every worksheet with Aspose.Cells for .NET
// Description: Loads an existing Excel file, reads a SVG logo into a byte array, and uses Aspose.Cells' AddSvg method to place the image at row 0, column 10 on each worksheet. The picture is set to free‑floating placement and the workbook is saved with the logo on all sheets.
// Keywords: Aspose.Cells | C# | AddSvg | SVG logo | batch insert image | top right placement | worksheet shape | free floating placement | Excel automation | multi‑sheet logo
// Common Searches: Aspose.Cells add the same SVG to every sheet | C# insert logo at top right of all worksheets | How to use AddSvg in Aspose.Cells .NET | Place free floating SVG picture in Excel with Aspose | Batch image insertion across multiple worksheets Aspose.Cells | Set SVG size and position in Aspose.Cells workbook
// Developer Intent: Add a single SVG logo to the top‑right corner of each worksheet in an existing Excel workbook.
// Use Cases: Brand all generated reports with a corporate SVG logo on every sheet. | Create a reusable template that automatically adds a logo to new worksheets. | Ensure consistent branding in multi‑sheet financial models or dashboards.
// AI Prompts: Write C# code using Aspose.Cells to insert an SVG image at column K, row 1 of every worksheet and save the file. | Explain how to control SVG width, height, and scaling when adding it to multiple worksheets with Aspose.Cells for .NET. | Provide best‑practice error handling for loading SVG data and inserting it into all worksheets using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an existing Excel file, reads a SVG logo into a byte array, and uses Aspose.Cells' AddSvg method to place the image at row 0, column 10 on each worksheet. The picture is set to free‑floating placement and the workbook is saved with the logo on all sheets.
class InsertSvgLogo
{
    static void Main()
    {
        try
        {
            string workbookPath = "input.xlsx";
            string svgPath = "company_logo.svg";
            string outputPath = "output.xlsx";

            // Ensure the workbook file exists before loading
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Ensure the SVG file exists before reading
            if (!File.Exists(svgPath))
            {
                Console.WriteLine($"SVG file not found: {svgPath}");
                return;
            }

            // Load SVG data into a byte array
            byte[] svgBytes = File.ReadAllBytes(svgPath);

            // Insert the SVG into each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                ShapeCollection shapes = sheet.Shapes;

                // Add SVG picture at specified position
                Picture picture = shapes.AddSvg(
                    topRow: 0,
                    top: 0,
                    leftColumn: 10,
                    left: 0,
                    height: -1,
                    width: -1,
                    svgData: svgBytes,
                    compatibleImageData: null);

                // Optional: set placement to free floating
                picture.Placement = PlacementType.FreeFloating;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
