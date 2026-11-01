// Title: Adjust WordArt shape fill transparency to 30% in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load a workbook, locate the first WordArt object, assign 0.3 to its Fill.Transparency, and save the file with Aspose.Cells in C#. | Iterate through worksheet.Shapes, identify shapes where IsWordArt is true, change the shape's fill opacity to 30 %, then write the updated workbook using Aspose.Cells.
// Common Searches: Aspose.Cells .NET example for changing WordArt opacity in Excel | C# code to make a WordArt object semi‑transparent with Aspose.Cells | How to adjust fill transparency of a WordArt shape using Aspose.Cells library | Set 30 percent transparency on Excel WordArt via C# Aspose.Cells
// Tags: Aspose.Cells shape fill transparency | C# WordArt opacity modification | Excel WordArt transparency example | set fill transparency Aspose.Cells .NET | programmatic shape opacity Excel C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads Input.xlsx, searches the first worksheet for a WordArt shape, sets its Fill.Transparency property to 0.3 (30 % transparency), saves the modified workbook as Output.xlsx, and prints status messages.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the WordArt shape
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Find the WordArt shape and set its transparency to 30%
            bool wordArtFound = false;
            foreach (Shape shape in worksheet.Shapes)
            {
                // Use the IsWordArt property to identify WordArt shapes
                if (shape.IsWordArt)
                {
                    // Transparency value is between 0 (opaque) and 1 (fully transparent)
                    shape.Fill.Transparency = 0.3; // 30% transparency
                    wordArtFound = true;
                    break; // Assuming only one WordArt needs adjustment
                }
            }

            if (!wordArtFound)
            {
                Console.WriteLine("No WordArt shape found in the worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
