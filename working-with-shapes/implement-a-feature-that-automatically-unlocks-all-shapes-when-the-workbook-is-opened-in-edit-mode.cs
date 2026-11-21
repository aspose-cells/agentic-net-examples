// Title: Automatically unlock all shapes in an Excel workbook on opening using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# unlock all shapes when workbook is opened | set IsLocked false for every shape in Excel file using Aspose.Cells | batch unlock worksheet shapes on workbook load in .NET | how to programmatically unlock shapes in all worksheets with Aspose.Cells | remove shape protection from Excel workbook on open using C#
// Tags: unlock shapes Aspose.Cells workbook | shape IsLocked property .NET | batch shape unlock Excel Aspose | auto unlock shapes on workbook load | iterate worksheet shapes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Loads an existing Excel file, iterates through each worksheet and its shapes, sets each shape's IsLocked property to false, and saves the workbook, ensuring all shapes are unlocked automatically when the file is opened.
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

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets and unlock each shape
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Shape shape in sheet.Shapes)
                    {
                        shape.IsLocked = false;
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
