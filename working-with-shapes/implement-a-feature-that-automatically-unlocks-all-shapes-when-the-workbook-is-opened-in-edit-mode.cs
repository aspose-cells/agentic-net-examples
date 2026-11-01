// Title: Automatically unlock all shapes in an Excel workbook on load using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, iterates through every worksheet and shape, sets each shape's IsLocked property to false, and saves the workbook. | Create a method that accepts a Workbook object and unlocks all shapes across all worksheets using the Aspose.Cells API. | Demonstrate how to configure Aspose.Cells so that shape protection is removed automatically when a workbook is opened in edit mode.
// Common Searches: asp.net unlock all shapes in Excel file using Aspose.Cells | how to remove shape protection from every worksheet with Aspose.Cells C# | set IsLocked false for all shapes in a workbook Aspose.Cells .NET | batch unlock Excel shapes programmatically Aspose.Cells example | unlock shapes on workbook open Aspose.Cells C#
// Tags: unlock shapes Aspose.Cells C# | shape IsLocked property Aspose.Cells | batch unlock Excel shapes .NET | iterate worksheets shapes Aspose.Cells | remove shape protection programmatically

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an existing Excel workbook with Aspose.Cells, iterates through each worksheet and every shape, sets the IsLocked flag of each shape to false, and then saves the modified workbook, effectively unlocking all shapes when the file is opened in edit mode.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
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
