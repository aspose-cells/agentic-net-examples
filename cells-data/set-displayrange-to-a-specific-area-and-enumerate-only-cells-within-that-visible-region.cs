// Title: Define a display range and iterate only the visible cells using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set visible range B2:D5 and loop through cells | How to enumerate only cells inside a defined range with Aspose.Cells .NET | Using Worksheet.Cells.CreateRange to iterate visible cells in C# | C# example for limiting cell iteration to a display range in Aspose.Cells | Iterate cells in a specific worksheet region with Aspose.Cells for .NET
// Tags: Aspose.Cells CreateRange cell enumeration | C# define worksheet display range Aspose.Cells | iterate cells within visible range Aspose.Cells | save workbook to .xlsx using Aspose.Cells | display range B2:D5 Aspose.Cells example

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// The sample creates a new workbook, defines a display range covering cells B2 to D5, obtains that range via Worksheet.Cells.CreateRange, iterates over each Cell in the range printing its name and value, and saves the workbook as 'DisplayRangeDemo.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the display range (example: B2:D5)
            int startRow = 1;      // B2 -> row index 1
            int startColumn = 1;   // B2 -> column index 1
            int endRow = 4;        // D5 -> row index 4
            int endColumn = 3;     // D5 -> column index 3

            // Calculate size of the range
            int totalRows = endRow - startRow + 1;
            int totalColumns = endColumn - startColumn + 1;

            // Retrieve the same range as an Aspose.Cells.Range object
            AsposeRange visibleRange = sheet.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);

            // Enumerate only cells within the visible region
            foreach (Cell cell in visibleRange)
            {
                Console.WriteLine($"Cell {cell.Name} = {cell.Value}");
            }

            // Save the workbook (optional)
            string outputPath = "DisplayRangeDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
