// Title: Log shapes anchored to error cells in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that scans all shapes on a worksheet, identifies those whose UpperLeftRow/Column points to a cell with CellValueType.IsError, and writes each shape name and cell address to the console. | Create a C# method that returns a List<(string ShapeName, string CellAddress)> containing every shape linked to an error cell in a workbook using Aspose.Cells. | Modify the shape‑iteration loop to capture error‑linked shapes and export the results to a text log file, handling any shape‑processing exceptions gracefully.
// Common Searches: asp.net c# detect shapes linked to #N/A cells with Aspose.Cells | Aspose.Cells iterate worksheet shapes and find error values in linked cells | C# code to log shape names that are anchored to error cells in Excel using Aspose | filter Excel shapes by cell error type Aspose.Cells .NET example | retrieve shapes pointing to cells with #DIV/0! using Aspose.Cells
// Tags: iterate worksheet shapes Aspose.Cells | identify shapes linked to error cells C# | export shape‑cell error mapping Aspose.Cells | shape anchor error detection .NET | filter shapes by CellValueType.IsError

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, iterates through all shapes on the first worksheet, checks the cell anchored at each shape's upper‑left corner for an error value (CellValueType.IsError), logs the shape name and cell reference when an error is found, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes in the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                try
                {
                    // Get the cell to which the shape is anchored (top‑left corner)
                    int row = shape.UpperLeftRow;
                    int column = shape.UpperLeftColumn;
                    Cell linkedCell = worksheet.Cells[row, column];

                    // Check if the linked cell contains an error value
                    if (linkedCell.Type == CellValueType.IsError)
                    {
                        // Log the shape name and the problematic cell reference
                        Console.WriteLine($"Shape '{shape.Name}' is linked to error cell {linkedCell.Name}.");
                    }
                }
                catch (Exception exShape)
                {
                    // Handle any unexpected errors while processing a shape
                    Console.WriteLine($"Error processing shape '{shape.Name}': {exShape.Message}");
                }
            }

            // Save the workbook (optional, depending on further processing)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
