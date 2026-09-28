// Title: Automatically unlock every shape in an Excel workbook when opened with Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to load an Excel file, loop through all worksheets and their Shapes collections, set each shape's IsLocked property to false, and save the workbook. | Update an existing Aspose.Cells .NET program so that it automatically removes shape protection for all shapes each time the workbook is opened in edit mode.
// Common Searches: asp.net unlock all shapes in Excel workbook using Aspose.Cells | c# Aspose.Cells set shape IsLocked false for every worksheet | how to programmatically remove shape protection from an Excel file with Aspose.Cells | batch unlock shapes in Excel file Aspose.Cells example
// Tags: unlock all shapes Aspose.Cells | set shape IsLocked property C# | iterate worksheets shapes Aspose.Cells | remove shape protection Excel .NET | batch shape unlocking Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace UnlockShapesExample
{
    // The example loads an existing Excel workbook, iterates through each worksheet and its Shapes collection, sets the IsLocked flag of every shape to false, and saves the modified file, ensuring all shapes are unlocked when the workbook is opened.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output_unlocked.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Unlock all shapes in each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Shape shape in sheet.Shapes)
                    {
                        shape.IsLocked = false;
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
