// Title: How to delete a TextBox shape from an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove textbox shape from worksheet | delete specific textbox in Excel file with Aspose.Cells .NET | programmatically remove shape from Excel sheet using Aspose.Cells | C# code to iterate and delete textbox shape in workbook | Aspose.Cells remove textbox without affecting other shapes
// Tags: Aspose.Cells Shapes.RemoveAt usage | C# reverse iterate worksheet shapes | Excel textbox deletion Aspose.Cells API | Aspose.Cells shape management .NET | preserve other shapes when deleting Excel shape

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, accesses a worksheet, iterates through its Shapes collection in reverse order, identifies TextBox shapes by type, removes the first matching textbox using the Shapes.RemoveAt method, and saves the modified workbook, handling file existence checks and error logging.
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
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Access the target worksheet (by name or index)
                Worksheet sheet = workbook.Worksheets["Sheet1"]; // or workbook.Worksheets[0]

                // Iterate through shapes in reverse order to safely remove items
                for (int i = sheet.Shapes.Count - 1; i >= 0; i--)
                {
                    try
                    {
                        Shape shape = sheet.Shapes[i];

                        // Identify textbox shapes by their type name (avoids direct enum reference)
                        if (shape.Type.ToString() == "TextBox")
                        {
                            // Remove the textbox shape from the worksheet
                            sheet.Shapes.RemoveAt(i);
                            // If only one textbox is expected, exit the loop
                            break;
                        }
                    }
                    catch (Exception exShape)
                    {
                        // Log shape-specific errors and continue processing other shapes
                        Console.WriteLine($"Error processing shape at index {i}: {exShape.Message}");
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook to the desired output path
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
