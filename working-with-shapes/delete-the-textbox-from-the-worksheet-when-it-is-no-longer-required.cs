// Title: Delete all TextBox shapes from an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# code to remove textbox shapes from an Excel worksheet with Aspose.Cells | How to iterate and delete specific shape types in an Aspose.Cells workbook | Aspose.Cells remove all TextBox objects from a sheet before saving | Delete shapes by type in Aspose.Cells .NET example | Remove unwanted text boxes from an Excel file using the Aspose.Cells API
// Tags: Aspose.Cells delete textbox shapes | C# iterate worksheet shapes collection | remove specific shape type Aspose.Cells | save workbook after shape removal Aspose.Cells | Aspose.Cells shape removal example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The sample loads an existing Excel file with Aspose.Cells, iterates the worksheet's Shapes collection in reverse, checks each shape's Type for "TextBox", removes matching shapes, and saves the workbook to a new file, handling missing files and shape‑processing errors.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputFile = "input.xlsx";
                const string outputFile = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(inputFile);

                // Get the target worksheet (adjust index or name as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Delete all textbox shapes from the worksheet
                // Iterate backwards because removing modifies the collection
                for (int i = worksheet.Shapes.Count - 1; i >= 0; i--)
                {
                    try
                    {
                        Shape shape = worksheet.Shapes[i];
                        // Use string comparison to avoid direct enum dependency
                        if (shape.Type.ToString() == "TextBox")
                        {
                            worksheet.Shapes.RemoveAt(i);
                        }
                    }
                    catch (Exception exShape)
                    {
                        Console.WriteLine($"Error processing shape at index {i}: {exShape.Message}");
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook (lifecycle rule: save)
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved to {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
