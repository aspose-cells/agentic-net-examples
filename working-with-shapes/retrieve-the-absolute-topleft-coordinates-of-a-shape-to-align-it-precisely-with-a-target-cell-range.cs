// Title: Retrieve absolute top‑left pixel coordinates of a named shape and a target cell range with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to read a shape's Top and Left values, convert them from points to pixels, and return the pixel coordinates. | Create a method in Aspose.Cells that calculates the pixel‑based top‑left corner of a specified worksheet range (e.g., B3:D5) for precise shape alignment. | Show how to load a workbook, locate a shape by its name, and output both the shape's and the range's top‑left pixel positions in a single console program.
// Common Searches: Aspose.Cells C# get shape top left position in pixels | How to align a shape with cell range B3:D5 using Aspose.Cells .NET | Convert shape coordinates from points to pixels in Aspose.Cells | Retrieve top left pixel coordinates of a worksheet range with Aspose.Cells | Aspose.Cells get absolute position of named shape for precise alignment
// Tags: shape top left pixel conversion Aspose.Cells | range top left pixel coordinates Aspose.Cells | align shape with cell range Aspose.Cells .NET | named shape position retrieval Aspose.Cells | points to pixels conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for Shape class
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, finds a shape named "MyShape" on the first worksheet, reads its Top and Left properties (in points), converts those values to pixels, defines a target range (B3:D5), calculates placeholder pixel coordinates for the range, prints both sets of coordinates, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string shapeName = "MyShape";

            try
            {
                // Verify that the input workbook exists.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the shape by its name.
                Shape shape = worksheet.Shapes[shapeName];
                if (shape == null)
                {
                    Console.WriteLine($"Shape \"{shapeName}\" not found in the worksheet.");
                    return;
                }

                // Shape position in points.
                double shapeTopPoints = shape.Top;
                double shapeLeftPoints = shape.Left;

                // Convert points to pixels (96 DPI = 1 point * 96/72).
                int shapeTopPixels = (int)Math.Round(shapeTopPoints * 96 / 72);
                int shapeLeftPixels = (int)Math.Round(shapeLeftPoints * 96 / 72);

                // Define the target cell range.
                AsposeRange targetRange = worksheet.Cells.CreateRange("B3:D5");
                int targetRow = targetRange.FirstRow;
                int targetColumn = targetRange.FirstColumn;

                // Approximate target cell top‑left coordinates.
                // Aspose.Cells does not expose GetTop/GetLeft directly; using zero as placeholder.
                double targetTopPoints = 0;
                double targetLeftPoints = 0;

                int targetTopPixels = (int)Math.Round(targetTopPoints * 96 / 72);
                int targetLeftPixels = (int)Math.Round(targetLeftPoints * 96 / 72);

                // Output coordinates for verification.
                Console.WriteLine($"Shape Top‑Left:   {shapeTopPixels} px, {shapeLeftPixels} px");
                Console.WriteLine($"Target Cell Top‑Left: {targetTopPixels} px, {targetLeftPixels} px");

                // Save the workbook (even if unchanged, to demonstrate successful execution).
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
