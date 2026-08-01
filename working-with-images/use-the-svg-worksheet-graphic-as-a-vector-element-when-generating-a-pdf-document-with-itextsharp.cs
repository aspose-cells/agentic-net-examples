// Title: Add an SVG shape to an Aspose.Cells worksheet and export it as a vector PDF (C#)
// Description: Loads an SVG file, inserts it into the first worksheet with the AddSvg method, saves the workbook as XLSX, then creates a PDF where the SVG remains a true vector graphic.
// Keywords: Aspose.Cells SVG | AddSvg C# | vector PDF export | SaveFormat.Pdf | Excel to PDF SVG | preserve vector graphics | C# worksheet image | SVG shape worksheet
// Common Searches: how to embed SVG in Aspose.Cells worksheet | export SVG as vector in PDF using Aspose.Cells | AddSvg method example C# | keep SVG vector when saving Excel as PDF | Aspose.Cells convert SVG to PDF without rasterizing
// Developer Intent: Insert an SVG graphic into a worksheet and generate a PDF that retains the SVG as a vector element.
// Use Cases: Include a company logo in Excel reports and keep it crisp in the exported PDF. | Embed engineering diagrams as SVG in spreadsheets for high‑quality PDF documentation. | Automate batch conversion of Excel templates with SVG icons to vector PDFs.
// AI Prompts: Show how to place the SVG using A1‑style cell references instead of zero‑based indexes. | Generate code that adds every SVG file from a folder to a worksheet and exports a single vector PDF. | Explain PDF rendering options for SVG in Aspose.Cells to control quality and file size.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an SVG file, inserts it into the first worksheet with the AddSvg method, saves the workbook as XLSX, then creates a PDF where the SVG remains a true vector graphic.
class Program
{
    static void Main()
    {
        try
        {
            const string svgPath = "graphic.svg";

            // Verify that the SVG file exists before attempting to read it
            if (!File.Exists(svgPath))
            {
                Console.WriteLine($"SVG file not found: {svgPath}");
                return;
            }

            // Load SVG file bytes (the SVG that will be used as a vector shape)
            byte[] svgData = File.ReadAllBytes(svgPath);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            ShapeCollection shapes = worksheet.Shapes;

            // Add the SVG to the worksheet at row 5, column 1 (zero‑based indexes)
            // Height and width set to -1 to keep the original SVG dimensions
            shapes.AddSvg(4, 0, 3, 0, -1, -1, svgData, null);

            // Save the workbook that now contains the SVG shape
            workbook.Save("WorkbookWithSvg.xlsx");

            // Save the same workbook as PDF; the SVG shape is preserved as a vector element
            workbook.Save("OutputWithSvg.pdf", SaveFormat.Pdf);

            Console.WriteLine("Workbook and PDF generated with SVG as a vector element.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
