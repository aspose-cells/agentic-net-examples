// Title: Set a dashed 2‑point line style for the 'Divider' shape in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, finds the shape named 'Divider' on the first worksheet, and changes its Line.DashStyle to Dash with a weight of 2 points using Aspose.Cells. | Generate a script that updates the border of a specific shape in Excel to a dashed line of 2‑point thickness via the Aspose.Cells Shape.Line properties. | Create a method that searches a worksheet for a shape by name and applies a dashed line style and 2‑point weight using the Aspose.Cells Drawing API in C#.
// Common Searches: Aspose.Cells C# change shape line dash style to dashed | How to set line weight of a specific shape in Excel with Aspose.Cells | C# code to modify border thickness of a named shape in a workbook | Find shape by name and update its line properties using Aspose.Cells .NET
// Tags: Aspose.Cells modify shape line style | Aspose.Cells set shape dash pattern | C# adjust Excel shape border thickness | Aspose.Cells Drawing API line properties

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, locates the shape called 'Divider' on the first worksheet, sets its line dash style to dashed and its weight to 2 points, and then saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException($"The input file '{inputPath}' was not found.");
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Locate the shape named "Divider"
                Shape dividerShape = null;
                foreach (Shape shape in worksheet.Shapes)
                {
                    if (shape.Name == "Divider")
                    {
                        dividerShape = shape;
                        break;
                    }
                }

                // If the shape is found, modify its line style and thickness
                if (dividerShape != null)
                {
                    // Set the line dash style to dashed
                    dividerShape.Line.DashStyle = MsoLineDashStyle.Dash;

                    // Set the line weight (thickness) to 2 points
                    dividerShape.Line.Weight = 2.0;
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log or display the exception details
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
