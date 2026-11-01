// Title: C# example: bring a named shape to the front layer in an Excel worksheet using Aspose.Cells
// Common Searches: Aspose.Cells C# set shape ZOrderPosition to bring shape forward | How to change Z-order of a specific shape in an Excel file with Aspose.Cells | Move an Excel shape to the front layer programmatically using Aspose.Cells .NET | Retrieve a shape by name and adjust its Z-order in an Aspose.Cells workbook | C# code to bring a chart shape to the front in an Excel worksheet with Aspose.Cells
// Tags: Aspose.Cells shape ZOrderPosition | C# Excel shape front layer | Aspose.Cells shape ordering API | Modify Excel shape Z-order with Aspose.Cells | Set shape Z-order position in worksheet

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // The sample loads an existing Excel workbook, accesses the "Sheet1" worksheet, retrieves a shape by its name, sets its ZOrderPosition to 0 to place the shape on the front layer, and saves the modified workbook to the specified output path.
    class Program
    {
        static void Main(string[] args)
        {
            // Define file paths and shape identifier (replace with actual values)
            string inputFilePath = "{InputFilePath}";
            string outputFilePath = "{OutputFilePath}";
            string shapeName = "{ShapeName}";
            // int shapeIndex = 0; // alternative if you prefer index

            try
            {
                // Ensure the input workbook exists
                if (!File.Exists(inputFilePath))
                {
                    Console.WriteLine($"Input file not found: {inputFilePath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputFilePath);

                // Access the target worksheet (by name or index)
                Worksheet worksheet = workbook.Worksheets["Sheet1"];
                if (worksheet == null)
                {
                    Console.WriteLine("Worksheet 'Sheet1' not found.");
                    return;
                }

                // Retrieve the shape by name (or use index if needed)
                Shape shape = worksheet.Shapes[shapeName];
                // Shape shape = worksheet.Shapes[shapeIndex]; // alternative

                if (shape == null)
                {
                    Console.WriteLine($"Shape '{shapeName}' not found in worksheet.");
                    return;
                }

                // Bring the shape to the front layer by setting its Z-order position
                // Lower ZOrderPosition values are rendered in front of higher values
                shape.ZOrderPosition = 0;

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputFilePath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputFilePath);
                Console.WriteLine($"Workbook saved successfully to {outputFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
