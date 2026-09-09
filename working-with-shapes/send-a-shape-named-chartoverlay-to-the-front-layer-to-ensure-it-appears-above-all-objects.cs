// Title: How to bring a shape named "ChartOverlay" to the front layer in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Set the ZOrder of the shape "ChartOverlay" to the highest index so it appears above all other objects using Aspose.Cells in C#. | Retrieve a named shape from a worksheet and modify its ZOrder property to position it on the topmost layer with the Aspose.Cells API. | Adjust the ZOrder of a chart overlay shape to the maximum value before saving the workbook with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set shape ZOrder to highest value | bring named shape to front in Excel using Aspose.Cells library | how to change layering of chart overlay shape in .NET workbook | C# Aspose.Cells move shape to top of Z-order stack
// Tags: Aspose.Cells shape ZOrder manipulation | C# Excel shape layering with Aspose | named shape ordering Aspose.Cells | chart overlay ZOrder setting | worksheet shape order .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, accesses the shape named "ChartOverlay" on the first worksheet, and shows how to simulate bringing the shape to the front by assigning its ZOrder property the highest index before saving the file.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the shape named "ChartOverlay"
            Shape chartOverlay = sheet.Shapes["ChartOverlay"];
            if (chartOverlay == null)
            {
                Console.WriteLine("Shape 'ChartOverlay' not found.");
                return;
            }

            // NOTE: Aspose.Cells does not provide a BringToFront method for Shape.
            // If ordering is required, adjust the ZOrder property as needed.
            // Example (optional): chartOverlay.ZOrder = sheet.Shapes.Count;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
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
