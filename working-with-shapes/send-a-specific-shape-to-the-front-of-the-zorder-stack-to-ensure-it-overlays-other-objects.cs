// Title: How to bring a specific shape to the front of the Z‑order stack in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Retrieve a shape by its index or name and set its ZOrderPosition to the highest value in an Aspose.Cells workbook. | Programmatically move an Excel shape to the top layer and save the updated file with Aspose.Cells in C#. | Adjust the Z‑order of a worksheet shape so it overlays all other objects using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# bring shape to front of worksheet | set ZOrderPosition for a shape in Aspose.Cells workbook | move Excel shape to top layer programmatically using .NET | how to change shape stacking order in an existing Excel file with Aspose.Cells | retrieve shape by index and change Z‑order Aspose.Cells C#
// Tags: Aspose.Cells set shape ZOrderPosition | C# bring Excel shape to front | modify shape stacking order Aspose.Cells | retrieve worksheet shape by index Aspose.Cells | save workbook after shape reorder Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for Shape class

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, accesses the first worksheet, obtains a shape by its index, sets the shape's ZOrderPosition to the maximum index so it appears above all other objects, ensures the output directory exists, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index or name as needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Retrieve the shape you want to bring to the front.
                // This example uses the shape's index; you can also use its name.
                int shapeIndex = 0; // replace with the actual index of the target shape
                if (shapeIndex < 0 || shapeIndex >= sheet.Shapes.Count)
                {
                    Console.WriteLine("Error: Shape index is out of range.");
                    return;
                }

                Shape shape = sheet.Shapes[shapeIndex];

                // Bring the shape to the front by setting its Z‑order position to the highest value
                shape.ZOrderPosition = sheet.Shapes.Count - 1;

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
