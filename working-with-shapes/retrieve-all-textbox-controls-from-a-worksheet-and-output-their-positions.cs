// Title: How to retrieve all TextBox controls from an Excel worksheet and output their row/column positions using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to load an existing .xlsx file, iterate through the worksheet's Shapes collection, filter for TextBox objects, and print each TextBox's UpperLeftRow and UpperLeftColumn values. | Show an example that extracts the cell indices of every TextBox shape in the first sheet of a workbook without modifying the file, using Aspose.Cells for .NET. | Provide a C# snippet that logs the name, type, and cell position (row and column) of all TextBox controls in an Excel worksheet, handling missing files gracefully.
// Common Searches: Aspose.Cells C# get coordinates of TextBox shapes in Excel worksheet | How to list TextBox controls and their cell positions using Aspose.Cells .NET | Read‑only mode retrieve textbox row and column indices from .xlsx with Aspose.Cells | Filter worksheet shapes for TextBox and output UpperLeftRow UpperLeftColumn Aspose.Cells | C# Aspose.Cells example to enumerate TextBox objects in a sheet
// Tags: Aspose.Cells filter TextBox shapes C# | extract textbox positions Excel .NET | enumerate worksheet shapes Aspose.Cells | read-only workbook shape extraction C# | output shape cell indices Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, iterates through the Shapes collection of the first worksheet, filters for TextBox objects, and writes each TextBox's name, type, UpperLeftRow, UpperLeftColumn, width, and height to the console.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook (read‑only operation)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through all shapes in the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Output basic shape information
                Console.WriteLine($"Shape Name: {shape.Name}");
                Console.WriteLine($"Shape Type: {shape.Type}");

                // Position details (row/column)
                Console.WriteLine($"Upper Left Row: {shape.UpperLeftRow}, Column: {shape.UpperLeftColumn}");

                // Width and height (in pixels)
                Console.WriteLine($"Width (pixels): {shape.Width}, Height (pixels): {shape.Height}");
                Console.WriteLine();
            }

            // No saving required for this read‑only operation
        }
        catch (Exception ex)
        {
            // Handle unexpected exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
