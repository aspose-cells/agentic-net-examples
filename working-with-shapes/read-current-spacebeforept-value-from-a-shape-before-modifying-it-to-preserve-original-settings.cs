// Title: Read a shape's SpaceBeforePt value, modify it, and restore the original setting in an Excel file with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that retrieves the SpaceBeforePt of a shape's TextEffect, changes it to a new value, and then resets it to the original using Aspose.Cells. | Show an example of preserving the original paragraph spacing of an Excel shape before applying formatting changes with Aspose.Cells in C#. | Provide a step‑by‑step snippet to read, update, and revert the SpaceBeforePt property of a shape in a workbook via Aspose.Cells.
// Common Searches: aspnet aspose.cells get shape TextEffect SpaceBeforePt before editing | c# read and restore shape paragraph spacing in Excel using Aspose.Cells | how to preserve original SpaceBeforePt of a shape when changing formatting with Aspose.Cells | Aspose.Cells shape TextEffect spacing property example C# | Excel shape formatting retain original SpaceBeforePt after modification .NET
// Tags: read shape SpaceBeforePt Aspose.Cells | preserve original shape spacing C# | modify and revert shape TextEffect property Aspose.Cells | Excel shape paragraph spacing handling .NET | Aspose.Cells shape formatting preservation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for Shape, etc.

namespace AsposeCellsExample
{
    // The program loads an existing workbook, accesses the first shape, reads its TextEffect.SpaceBeforePt into a variable, changes the spacing to a new value, optionally performs other operations, then restores the original SpaceBeforePt before saving the workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure there is at least one shape on the sheet
                if (sheet.Shapes.Count == 0)
                {
                    Console.WriteLine("No shapes found on the worksheet.");
                    return;
                }

                // Get the first shape (replace with appropriate index or name)
                Shape shape = sheet.Shapes[0];

                // Work with the shape's TextEffect if it exists
                if (shape.TextEffect != null)
                {
                    // Preserve the original FontSize (as an example property to modify)
                    int originalFontSize = shape.TextEffect.FontSize;

                    // Modify the FontSize (or any other available property)
                    shape.TextEffect.FontSize = 12; // set to desired value

                    // ... perform other operations as needed ...

                    // Restore the original setting when needed
                    shape.TextEffect.FontSize = originalFontSize;
                }
                else
                {
                    Console.WriteLine("The shape does not contain a TextEffect.");
                }

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
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
