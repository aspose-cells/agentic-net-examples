// Title: Add SVG as a Vector Shape in Aspose.Cells and Export to PDF (C#)
// Description: C# example that loads an SVG file, inserts it as a vector shape on a worksheet using Aspose.Cells AddSvg, and saves the workbook directly to PDF while preserving the SVG’s vector quality. Includes file‑existence check and error handling.
// Keywords: Aspose.Cells AddSvg | C# SVG to PDF | vector graphics PDF Aspose | export worksheet as PDF | preserve SVG vector quality | Aspose.Cells PDF export | embed SVG in spreadsheet
// Common Searches: Aspose.Cells add SVG shape C# | export worksheet with SVG to PDF | keep SVG vector when converting to PDF Aspose | C# example AddSvg method | how to embed SVG in Aspose.Cells workbook
// Developer Intent: Insert an SVG file as a vector graphic into an Aspose.Cells workbook and generate a PDF that retains the SVG in vector form.
// Use Cases: Create printable reports with a company logo in SVG format that stays sharp in the PDF output. | Generate invoices that embed scalable SVG icons, ensuring crisp rendering on any device. | Automate batch conversion of spreadsheets containing SVG assets to PDF without rasterization.
// AI Prompts: Provide C# code to load an SVG, add it as a vector shape to an Aspose.Cells worksheet, and save the workbook as a PDF preserving vector data. | Show an Aspose.Cells AddSvg example with specific row and column coordinates and PDF export. | Explain how to verify an SVG file exists and ensure vector graphics are retained during PDF conversion with Aspose.Cells.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that loads an SVG file, inserts it as a vector shape on a worksheet using Aspose.Cells AddSvg, and saves the workbook directly to PDF while preserving the SVG’s vector quality. Includes file‑existence check and error handling.
class SvgToPdfWithVector
{
    static void Main()
    {
        try
        {
            // Path to the SVG file to be used as a worksheet graphic
            string svgPath = "image.svg";

            // Verify that the SVG file exists to avoid FileNotFoundException
            if (!File.Exists(svgPath))
            {
                Console.WriteLine($"SVG file not found: {Path.GetFullPath(svgPath)}");
                return;
            }

            // Load the SVG data into a byte array
            byte[] svgBytes = File.ReadAllBytes(svgPath);

            // -----------------------------------------------------------------
            // Step 1: Create a workbook and add the SVG as a vector shape.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            ShapeCollection shapes = sheet.Shapes;

            // Add the SVG to the worksheet (positioned at row 4, column 5, default size)
            // The compatibleImageData parameter is set to null because we only need the vector representation.
            shapes.AddSvg(
                topRow: 4,          // Upper‑left row index
                top: 0,             // Vertical offset in pixels
                leftColumn: 5,      // Upper‑left column index
                left: 0,            // Horizontal offset in pixels
                height: -1,         // Use default height
                width: -1,          // Use default width
                svgData: svgBytes,
                compatibleImageData: null);

            // -----------------------------------------------------------------
            // Step 2: Save the workbook directly to PDF.
            // Aspose.Cells preserves vector graphics (including SVG) when exporting to PDF.
            // -----------------------------------------------------------------
            string pdfOutputPath = "output.pdf";
            workbook.Save(pdfOutputPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF generated successfully at: {Path.GetFullPath(pdfOutputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
