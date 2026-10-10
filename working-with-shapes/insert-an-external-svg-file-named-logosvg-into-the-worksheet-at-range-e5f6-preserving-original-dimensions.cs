// Title: Insert an external SVG file (logo.svg) into cells E5:F6 of an Excel worksheet using Aspose.Cells for .NET while preserving its original dimensions
// AI Prompts: Insert logo.svg into the range E5:F6 of a new workbook and keep its original dimensions using Aspose.Cells C#. | Add an external SVG picture to specific cells, set Placement to MoveAndSize, and save the workbook with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# insert SVG into specific cell range E5:F6 | preserve original size of SVG when adding picture to Excel worksheet using Aspose.Cells | set picture placement to MoveAndSize for SVG image in Aspose.Cells .NET | how to add external logo.svg to Excel file with Aspose.Cells C# example
// Tags: pictures.add svg Aspose.Cells | svg insertion into worksheet range | moveandsize placement Aspose.Cells | c# add external image to Excel cells | preserve svg dimensions Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The code creates a new workbook, verifies that logo.svg exists, inserts the SVG into cell E5 (spanning E5:F6) using the Pictures.Add method, sets the picture's Placement to MoveAndSize to retain its original dimensions, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the external SVG file
            string svgPath = "logo.svg";

            // Verify that the SVG file exists before attempting to add it
            if (!File.Exists(svgPath))
            {
                Console.WriteLine($"SVG file not found: {svgPath}");
                return;
            }

            // Add the SVG picture to cell E5 (row index 4, column index 4)
            // Add method returns the picture index; retrieve the Picture object afterwards
            int pictureIndex = sheet.Pictures.Add(4, 4, svgPath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Make the picture move and size with cells
            picture.Placement = PlacementType.MoveAndSize;

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
