// Title: How to list SmartArt shape names in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get names of SmartArt objects in a worksheet | list all SmartArt diagrams from an Excel file using Aspose.Cells | iterate worksheet shapes and filter SmartArt in .NET | retrieve SmartArt shape collection from a workbook with Aspose.Cells | display SmartArt diagram names from the first worksheet using C#
// Tags: Aspose.Cells enumerate SmartArt shapes | C# list SmartArt names Excel | worksheet.Shapes iteration Aspose.Cells | filter SmartArt objects .NET | retrieve SmartArt diagram names

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, accesses the first worksheet, iterates through its Shapes collection, and prints the Name property of each shape, allowing you to identify and list SmartArt objects present in the sheet.
    class Program
    {
        static void Main()
        {
            const string inputFile = "input.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputFile);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Iterate through all shapes in the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Print the name of each shape (original intent was SmartArt)
                    Console.WriteLine(shape.Name);
                }
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
