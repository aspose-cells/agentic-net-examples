// Title: Copy cell formatting from A1:B5 to C1:D5 with Aspose.Cells for .NET
// Description: This example demonstrates how to load an Excel workbook (or create a placeholder if missing), define source and target ranges, transfer only the style information using the CopyStyle method, and save the result to a new file. No cell values are moved—only visual formatting is duplicated.
// Keywords: Aspose.Cells CopyStyle | C# Excel formatting copy | transfer cell style Aspose | range formatting Aspose.Cells | copy only style .NET
// Common Searches: Aspose.Cells copy only formatting between ranges | C# CopyStyle method example | how to duplicate Excel cell style with Aspose | copy style A1:B5 to C1:D5 Aspose.Cells | copy cell formatting without values .NET
// Developer Intent: Duplicate the visual style of cells A1:B5 into C1:D5 while leaving the underlying data unchanged.
// Use Cases: Apply a consistent theme to new worksheet sections without altering existing data. | Generate report templates where formatting is reused across multiple columns. | Refresh data in a sheet while preserving the original layout and styling.
// AI Prompts: Write C# code that uses Aspose.Cells to copy only the formatting from range A1:B5 to C1:D5 and saves the workbook as output.xlsx. | Explain which style attributes (fonts, borders, colors, etc.) are copied by the CopyStyle method and how it differs from the Copy method. | Provide a snippet that copies formatting, then clears the source cells' values while keeping their styles intact.

using System;
using System.IO;
using Aspose.Cells;

// This example demonstrates how to load an Excel workbook (or create a placeholder if missing), define source and target ranges, transfer only the style information using the CopyStyle method, and save the result to a new file. No cell values are moved—only visual formatting is duplicated.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook or create a new one if the file is missing
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                // Optional: add sample data to the source range
                ws.Cells["A1"].PutValue("Sample1");
                ws.Cells["B1"].PutValue("Sample2");
                // Save placeholder input for future runs
                workbook.Save(inputPath);
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Define source and destination ranges (use fully qualified Aspose.Cells.Range)
            Aspose.Cells.Range sourceRange = sheet.Cells.CreateRange("A1:B5");
            Aspose.Cells.Range destinationRange = sheet.Cells.CreateRange("C1:D5");

            // Copy only formatting (style) from source to destination
            destinationRange.CopyStyle(sourceRange);

            // Save the workbook with applied formatting changes
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
