// Title: Identify and log Excel shapes that sit on error cells using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells find shapes positioned on cells with error values | how to log Excel shapes that overlap error cells using Aspose.Cells | iterate worksheet shapes and check for CellValueType.IsError in Aspose.Cells | detect shapes linked to error cells in an Excel workbook with Aspose.Cells | Aspose.Cells shape UpperLeftRow UpperLeftColumn error detection example
// Tags: Aspose.Cells iterate worksheet shapes | detect error cells linked to shapes | log shapes over error cells in Excel | CellValueType.IsError shape position check | C# Aspose.Cells shape error detection

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook with Aspose.Cells, loops through all shapes on the first worksheet, obtains each shape's UpperLeftRow and UpperLeftColumn to locate the underlying cell, checks if the cell's type is an error using CellValueType.IsError, and writes a console message identifying shapes that sit on error cells before optionally saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                try
                {
                    // Retrieve the cell coordinates the shape is positioned at
                    int row = shape.UpperLeftRow;
                    int column = shape.UpperLeftColumn;

                    // Access the cell at that position
                    Cell linkedCell = worksheet.Cells[row, column];

                    // Determine if the cell contains an error value
                    if (linkedCell.Type == CellValueType.IsError)
                    {
                        // Log the problematic shape information
                        Console.WriteLine(
                            $"Shape '{shape.Name}' is positioned over error cell {linkedCell.Name}.");
                    }
                }
                catch (Exception exShape)
                {
                    // Handle any unexpected errors while processing a specific shape
                    Console.WriteLine($"Error processing shape '{shape.Name}': {exShape.Message}");
                }
            }

            // Save the workbook (optional, if any changes were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
