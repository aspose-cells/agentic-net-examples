// Title: How to send a worksheet shape to the back layer and position it at cell A1 using Aspose.Cells for .NET
// AI Prompts: Move the first shape on a worksheet to the lowest Z‑order and set its Left and Top properties to zero points with Aspose.Cells in C#. | Programmatically place a shape behind all other objects and align its top‑left corner with cell A1 in an Excel workbook using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# set shape behind other objects and align to cell A1 | C# code to change ZOrderPosition of an Excel shape using Aspose.Cells | How to position a worksheet shape at the top left corner with Aspose.Cells .NET | Move Excel shape to back layer programmatically with Aspose.Cells library
// Tags: Aspose.Cells shape ZOrderPosition | Aspose.Cells move shape to back layer | Aspose.Cells align shape to cell A1 | Aspose.Cells set shape absolute position | Aspose.Cells worksheet shape positioning

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing workbook, retrieves the first shape on the first worksheet, sets its ZOrderPosition to 0 to send it to the back layer, resets its Left and Top coordinates to 0 points to align with the top‑left corner of cell A1, and saves the modified file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or specify by name/index)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one shape in the worksheet
            if (worksheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the worksheet.");
                return;
            }

            // Get the shape you want to modify (first shape in the collection)
            Shape shape = worksheet.Shapes[0];

            try
            {
                // Move the shape to the back layer by setting its Z-order position to the lowest value
                shape.ZOrderPosition = 0;

                // Align the shape with the top‑left corner of the worksheet (cell A1)
                shape.Left = 0; // Horizontal position in points
                shape.Top = 0;  // Vertical position in points
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error modifying shape: {ex.Message}");
                return;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
