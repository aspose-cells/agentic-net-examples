// Title: How to bring a SmartArt shape to the front in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an existing .xlsx file with Aspose.Cells, finds the first SmartArt shape on a worksheet, sets its ZOrderPosition to the topmost layer, and saves the workbook. | Show a step‑by‑step example of iterating through worksheet.Shapes, detecting IsSmartArt, and changing the shape's Z‑order so it renders above all other objects. | Provide a reusable C# method that accepts a workbook and a shape identifier, then moves the specified SmartArt shape to the front using Aspose.Cells.
// Common Searches: Aspose.Cells C# set SmartArt ZOrderPosition to front of other shapes | How to change the layer order of a SmartArt object in an Excel file with Aspose.Cells | C# code to bring a specific shape to the top in an Excel worksheet using Aspose.Cells | Move SmartArt shape to front programmatically in .NET Excel library
// Tags: Aspose.Cells set SmartArt ZOrderPosition | C# move Excel shape to front | worksheet shape ordering Aspose.Cells | modify SmartArt layer order .NET | Excel shape Z‑order manipulation C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, scans the worksheet's Shapes collection for a SmartArt object, sets its ZOrderPosition to 0 (the frontmost layer), and saves the updated file.
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
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Identify the SmartArt shape using the IsSmartArt property
                if (shape.IsSmartArt)
                {
                    // Bring the SmartArt shape to the front by setting its Z‑order position
                    // Lower ZOrderPosition values are rendered in front of higher values
                    shape.ZOrderPosition = 0;

                    // If only one SmartArt needs to be processed, exit the loop
                    break;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log or display the exception details as needed
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
