// Title: Retrieve the Shapes collection from the active worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get shapes collection from the active worksheet | list all shape names and types in an Excel file with Aspose.Cells | how to enumerate worksheet shapes in a .xlsx using Aspose.Cells | access shapes on the current sheet in Aspose.Cells for .NET | iterate through Excel shapes with Aspose.Cells C# example
// Tags: Aspose.Cells retrieve worksheet shapes C# | enumerate shape collection Aspose.Cells .xlsx | list shape names and types Aspose.Cells | active worksheet shape extraction .NET | C# iterate Excel shapes Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, obtains the active worksheet, accesses its Shapes collection, and iterates through each Shape to output the shape's Name and Type, with file‑existence checking and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

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

            // Get the currently active worksheet
            Worksheet worksheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

            // Retrieve the collection of shapes on the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // List each shape's name and type
            foreach (Shape shape in shapes)
            {
                Console.WriteLine($"Shape Name: {shape.Name}, Type: {shape.Type}");
            }

            // If any modifications are made to shapes, you can save the workbook
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
