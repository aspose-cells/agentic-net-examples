// Title: Insert SVG into Excel range E5:F6 with original size using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to load an external SVG file (logo.svg) into a byte array and insert it as a shape at cells E5:F6 in a new workbook. The AddSvg method is called with width and height set to -1 to retain the SVG's native dimensions, then the workbook is saved as output.xlsx.
// Keywords: Aspose.Cells SVG insertion | C# AddSvg method | preserve SVG dimensions | insert SVG into Excel range | load SVG from file Aspose.Cells
// Common Searches: Aspose.Cells add SVG to specific cells | keep original size when inserting SVG in Excel C# | how to use AddSvg with range E5:F6 | load external SVG into Aspose.Cells workbook
// Developer Intent: Place an external SVG file into the E5:F6 cell block of an Excel worksheet while maintaining its original dimensions.
// Use Cases: Embedding a company logo SVG in a report template without distortion. | Adding vector icons to a dashboard worksheet for crisp scaling. | Programmatically inserting branded SVG graphics into generated spreadsheets.
// AI Prompts: Write C# code that uses Aspose.Cells to load a logo.svg file and insert it at range E5:F6, preserving its native size. | Show error‑handling best practices for reading an SVG file before adding it as a shape with Aspose.Cells. | Explain how to reposition or resize an SVG after insertion with the AddSvg method in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to load an external SVG file (logo.svg) into a byte array and insert it as a shape at cells E5:F6 in a new workbook. The AddSvg method is called with width and height set to -1 to retain the SVG's native dimensions, then the workbook is saved as output.xlsx.
class InsertSvg
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the SVG file
            string svgPath = "logo.svg";

            // Verify that the SVG file exists before attempting to load it
            if (!File.Exists(svgPath))
            {
                Console.WriteLine($"SVG file not found: {svgPath}");
                return;
            }

            // Load the SVG file into a byte array
            byte[] svgData;
            using (FileStream fs = new FileStream(svgPath, FileMode.Open, FileAccess.Read))
            {
                svgData = new byte[fs.Length];
                fs.Read(svgData, 0, svgData.Length);
            }

            // Insert the SVG at range E5:F6 (zero‑based indices: row 4, column 4)
            // Height and width set to -1 to preserve the original SVG dimensions
            ShapeCollection shapes = worksheet.Shapes;
            shapes.AddSvg(4, 4, 0, 0, -1, -1, svgData, null);

            // Save the workbook
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
