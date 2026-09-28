// Title: How to list all TextBox shapes in an Excel worksheet and display their cell positions using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, iterates through Worksheet.Shapes, selects only shapes whose type is TextBox, and prints each TextBox's name along with UpperLeftRow, UpperLeftColumn, LowerRightRow, and LowerRightColumn. | Extend the sample to also compute and output each TextBox's width and height in points while preserving the existing file‑existence check and per‑shape error handling.
// Common Searches: Aspose.Cells C# get coordinates of TextBox shapes in a worksheet | C# enumerate shapes in an Excel file and find TextBox locations using Aspose | How to retrieve row and column indices of a TextBox control with Aspose.Cells | List all TextBox objects and their cell ranges in an .xlsx using Aspose.Cells for .NET
// Tags: Aspose.Cells shape enumeration | Aspose.Cells TextBox position extraction | Aspose.Cells worksheet shape coordinates | C# Aspose.Cells retrieve shape cell range | Aspose.Cells .xlsx shape location

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook, iterates over all shapes on the first worksheet, filters shapes of type TextBox, and writes each TextBox's name together with its upper‑left and lower‑right row/column indices to the console, handling missing files and shape‑specific errors gracefully.
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes in the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                try
                {
                    // Identify TextBox controls without directly referencing ShapeType enum
                    if (shape.Type.ToString() == "TextBox")
                    {
                        // Retrieve position details (row/column indices)
                        int upperLeftRow = shape.UpperLeftRow;
                        int upperLeftColumn = shape.UpperLeftColumn;
                        int lowerRightRow = shape.LowerRightRow;
                        int lowerRightColumn = shape.LowerRightColumn;

                        // Output the TextBox name and its position
                        Console.WriteLine($"TextBox '{shape.Name}' Position:");
                        Console.WriteLine($"  Upper-Left: Row {upperLeftRow}, Column {upperLeftColumn}");
                        Console.WriteLine($"  Lower-Right: Row {lowerRightRow}, Column {lowerRightColumn}");
                    }
                }
                catch (Exception shapeEx)
                {
                    // Handle any shape-specific errors without stopping the whole process
                    Console.WriteLine($"Error processing shape '{shape.Name}': {shapeEx.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
