// Title: Change the linked cell of a CheckBox shape in an existing Excel workbook using Aspose.Cells for .NET
// AI Prompts: Load the workbook, locate the shape named "CheckBox 1" on the first worksheet, assign the linked cell address "C5" to the checkbox, and save the file with Aspose.Cells in C#. | Programmatically modify the cell reference a form‑control CheckBox points to by updating its LinkedCell attribute via the Aspose.Cells API.
// Common Searches: Aspose.Cells C# set linked cell for a checkbox shape | how to change the cell reference of an Excel form control checkbox using .NET | update checkbox linked cell address in an existing workbook with Aspose.Cells | modify form control checkbox cell link after layout change in Excel C#
// Tags: Aspose.Cells set CheckBox linked cell | C# update Excel form control cell reference | modify checkbox shape linked cell property | change linked cell address .NET Excel | update CheckBox LinkedCell Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, finds a CheckBox shape named "CheckBox 1" on the first worksheet, changes its LinkedCell to "C5", and saves the workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index or name as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Find the CheckBox shape by its name
                CheckBox checkBox = null;
                foreach (Shape shape in worksheet.Shapes)
                {
                    if (shape.Name == "CheckBox 1" && shape is CheckBox cb)
                    {
                        checkBox = cb;
                        break;
                    }
                }

                if (checkBox == null)
                {
                    Console.WriteLine("CheckBox named 'CheckBox 1' was not found.");
                    return;
                }

                // Update the linked cell reference to the new address (e.g., "C5")
                checkBox.LinkedCell = "C5";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the updated linked cell reference
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
