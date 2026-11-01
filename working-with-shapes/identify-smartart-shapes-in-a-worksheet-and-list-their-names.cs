// Title: How to find and list SmartArt shape names in each worksheet using Aspose.Cells for .NET
// Common Searches: C# code to enumerate SmartArt shapes in an Excel worksheet with Aspose.Cells | retrieve SmartArt object names from a workbook using Aspose.Cells .NET | list all SmartArt shapes across worksheets in an .xlsx file via Aspose.Cells | Aspose.Cells check if shape is SmartArt and get its name in C# | iterate through ShapeCollection to find SmartArt in Excel using Aspose.Cells
// Tags: Aspose.Cells SmartArt shape enumeration | IsSmartArt property usage | ShapeCollection iteration for SmartArt | extract SmartArt name in .NET | Excel workbook SmartArt identification

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an Excel workbook, loops through each worksheet, iterates the ShapeCollection, uses the IsSmartArt flag to detect SmartArt objects, and writes each SmartArt shape's Name together with its worksheet name to the console. The workbook is then saved unchanged.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of shapes on the worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Examine each shape for SmartArt
                foreach (Shape shape in shapes)
                {
                    if (shape.IsSmartArt)
                    {
                        Console.WriteLine($"Worksheet: {sheet.Name}, SmartArt Name: {shape.Name}");
                    }
                }
            }

            // Save the workbook (optional if no changes are made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
