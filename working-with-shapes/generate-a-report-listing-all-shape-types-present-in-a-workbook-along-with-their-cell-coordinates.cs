// Title: Create a C# console application that lists every shape’s type and its top‑left cell address for each worksheet in an Excel file using Aspose.Cells
// AI Prompts: Generate a console‑based report that iterates through all worksheets, reads each shape’s Name, Type, UpperLeftRow and UpperLeftColumn, converts the indices to an A1 address, and prints the information. | Write C# code with Aspose.Cells to enumerate shapes on every sheet, capture their type and cell location, and output a formatted table to the standard output. | Modify the sample to write the shape type and cell coordinate data into a CSV file instead of displaying it on the console.
// Common Searches: Aspose.Cells C# list shape types with cell addresses in each worksheet | How to get the top left cell of a shape using Aspose.Cells .NET | C# program to enumerate Excel shapes and output their positions | Export Aspose.Cells shape information to CSV in .NET
// Tags: Aspose.Cells enumerate worksheet shapes | Aspose.Cells get shape cell address | C# list Excel shape types | Aspose.Cells shape report console | Aspose.Cells export shape data to CSV

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Uses Aspose.Cells for .NET to load an Excel workbook, loop through each worksheet, retrieve every shape’s name, type, and upper‑left row/column, convert the indices to an A1 cell reference, and print a formatted report to the console (with optional CSV export).
class ShapeReport
{
    static void Main()
    {
        // Path to the input workbook
        string workbookPath = "input.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet: {sheet.Name}");

                // Get the collection of shapes on the current worksheet
                ShapeCollection shapes = sheet.Shapes;

                // If there are no shapes, indicate that and continue to next sheet
                if (shapes.Count == 0)
                {
                    Console.WriteLine("  No shapes found.");
                    continue;
                }

                // Enumerate each shape
                foreach (Shape shape in shapes)
                {
                    try
                    {
                        // Determine the shape type as a string
                        string shapeType = shape.Type.ToString();

                        // Retrieve the top‑left cell coordinates of the shape
                        int row = shape.UpperLeftRow;          // Upper-left row index (0‑based)
                        int column = shape.UpperLeftColumn;    // Upper-left column index (0‑based)
                        string cellAddress = CellsHelper.CellIndexToName(row, column);

                        // Output the shape information
                        Console.WriteLine($"  Shape Name: {shape.Name}");
                        Console.WriteLine($"    Type: {shapeType}");
                        Console.WriteLine($"    Position: {cellAddress}");
                    }
                    catch (Exception exShape)
                    {
                        // Handle errors related to a specific shape
                        Console.WriteLine($"  Error processing shape \"{shape.Name}\": {exShape.Message}");
                    }
                }
            }

            Console.WriteLine("Report generation completed.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }

        // Keep the console window open
        Console.ReadKey();
    }
}
