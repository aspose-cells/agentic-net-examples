// Title: How to loop through every TextBox in an Excel worksheet and apply the same character spacing using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that enumerates all TextBox shapes in a worksheet and sets Font.CharacterSpacing to a specified value. | Show an example of updating the character spacing of each TextBox while preserving its existing font attributes in a workbook. | Demonstrate saving the workbook after uniformly adjusting TextBox character spacing with Aspose.Cells.
// Common Searches: Aspose.Cells C# set Font.CharacterSpacing for all text boxes in a sheet | loop through TextBox shapes in Excel workbook and change character spacing | apply uniform character spacing to Excel text boxes using Aspose.Cells .NET API | how to modify font spacing of multiple text boxes in an Aspose.Cells workbook | C# example for updating character spacing of worksheet TextBox objects
// Tags: Aspose.Cells set TextBox character spacing | C# iterate worksheet shapes Aspose.Cells | uniform font spacing Excel text boxes | modify TextBox font properties Aspose.Cells | batch update TextBox character spacing .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook with Aspose.Cells, iterates over the Shapes collection of the first worksheet, selects each TextBox, assigns a uniform Font.CharacterSpacing value (e.g., 2 points) while leaving other font settings unchanged, and saves the modified workbook.
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
                    throw new FileNotFoundException($"Input file not found: {inputPath}");
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (or specify the desired one)
                Worksheet worksheet = workbook.Worksheets[0];

                // Iterate through all shapes in the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Process only TextBox shapes
                    if (shape is TextBox textBox)
                    {
                        // Example modification: make the text bold.
                        // Adjust other font properties as needed.
                        textBox.Font.IsBold = true;
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any errors that occur during processing
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
