// Title: Update a shape's connection point offset to (10,20) points and save the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook with Aspose.Cells, retrieve the first shape on the first worksheet, set Shape.Left = 10 and Shape.Top = 20 to represent a (10,20) offset, then save the workbook to a new file. | In C#, emulate a shape's connection site offset by adjusting its X and Y coordinates via the Left and Top properties, and persist the changes with Workbook.Save. | If the API version supports SetConnectionSiteOffset, call it with a PointF(10,20); otherwise achieve the same result by setting Left and Top, then write the workbook back to disk.
// Common Searches: how to change a shape's connection site offset in Aspose.Cells C# | move Excel shape to specific X Y coordinates using Aspose.Cells for .NET | C# code to update shape position and save workbook with Aspose.Cells | Aspose.Cells set shape left and top properties example | persist shape coordinate changes in an Excel file using Aspose.Cells
// Tags: Aspose.Cells shape position update C# | set shape left top Aspose.Cells .NET | modify Excel shape connection offset | save workbook after shape reposition Aspose.Cells | C# Aspose.Cells shape manipulation

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an existing Excel file, checks for at least one shape on the first worksheet, sets the shape's Left to 10 points and Top to 20 points (as a proxy for a (10,20) connection point offset), and then saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one shape in the worksheet
            if (worksheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the worksheet.");
                return;
            }

            // Retrieve the first shape
            Shape shape = worksheet.Shapes[0];

            // Update the position of the shape as an alternative to setting a connection site offset
            // (Aspose.Cells for .NET does not expose SetConnectionSiteOffset in this version)
            shape.Left = 10; // X position (in points)
            shape.Top = 20;  // Y position (in points)

            // Persist the changes to a new file (or overwrite the original)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
