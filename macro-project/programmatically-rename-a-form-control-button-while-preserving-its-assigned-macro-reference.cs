// Title: Programmatically rename an Excel form control button while keeping its macro reference using Aspose.Cells for .NET
// AI Prompts: Generate C# code that finds a shape named 'OldButtonName' on the first worksheet and changes its Name to 'NewButtonName' without altering the shape's Macro property. | Show how to iterate through worksheet shapes in Aspose.Cells to rename a specific button and preserve its assigned macro. | Provide a complete example that loads an .xlsx file, renames a form control button, and saves the workbook, ensuring the macro link remains intact.
// Common Searches: Aspose.Cells C# rename form control button without losing macro | how to change button name in Excel workbook using Aspose.Cells while keeping macro reference | preserve macro assignment when renaming shape in Aspose.Cells .NET | C# code to update Excel button name and retain its macro using Aspose.Cells | rename Excel form control shape programmatically with Aspose.Cells and keep macro link
// Tags: rename shape Aspose.Cells C# | preserve macro on shape rename Aspose.Cells | modify form control button name .xlsx Aspose.Cells | Aspose.Cells shape properties update | Excel macro link retention Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, searches the first worksheet for a shape named 'OldButtonName', changes its Name to 'NewButtonName' while the Macro property remains unchanged, and saves the workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Rename the button (or any shape) with the specified old name
                    if (shape.Name == "OldButtonName")
                    {
                        shape.Name = "NewButtonName";
                        // The assigned macro reference (shape.Macro) is preserved automatically
                        break; // Exit after renaming the target shape
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
                // Handle any runtime exceptions
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
