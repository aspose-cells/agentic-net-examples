// Title: Create a zero‑offset duplicate of an existing cell range in Aspose.Cells for .NET and apply a yellow background style
// AI Prompts: Generate a new Aspose.Cells range that references the same cells as an existing range without shifting its position, then set a solid yellow background. | Use the FirstRow, FirstColumn, RowCount, and ColumnCount properties of a range to clone it and apply a shading style in C#. | Call Worksheet.Cells.CreateRange with the original range's coordinates to create a duplicate range and style the duplicate cells.
// Common Searches: Aspose.Cells how to copy a range without moving it in C# | Create a duplicate range using original range coordinates Aspose.Cells .NET | Apply background color to a cloned range in Aspose.Cells | Zero offset range creation from existing range Aspose.Cells example | C# Aspose.Cells duplicate range and style cells
// Tags: duplicate range Aspose.Cells C# | zero offset range creation Aspose.Cells | apply background style to range Aspose.Cells | create range from coordinates Aspose.Cells | clone cell range .NET Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook, defines a range (A1:B2), creates a zero‑offset duplicate of that range using its coordinates, applies a solid yellow background style to the duplicate, and saves the modified workbook.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create the original range (example: A1:B2)
            AsposeRange originalRange = sheet.Cells.CreateRange("A1", "B2");

            // Create a duplicate reference to the same range using its coordinates
            AsposeRange duplicateRange = sheet.Cells.CreateRange(
                originalRange.FirstRow,
                originalRange.FirstColumn,
                originalRange.RowCount,
                originalRange.ColumnCount);

            // Example operation on the duplicate range: set background color to yellow
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.Yellow;
            style.Pattern = BackgroundType.Solid;

            StyleFlag flag = new StyleFlag
            {
                CellShading = true
            };

            duplicateRange.ApplyStyle(style, flag);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
