// Title: How to turn off row grand totals in an Excel pivot table using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide row grand total in existing pivot table | disable row grand totals programmatically with Aspose.Cells .NET | remove row grand total from Excel pivot table using Aspose.Cells API | set PivotTable.RowGrand to false in C# Aspose.Cells example | Excel pivot table row grand total suppression using Aspose.Cells library
// Tags: Aspose.Cells PivotTable.RowGrand property | C# disable row grand totals Excel | modify pivot table row totals Aspose.Cells | Excel workbook pivot table customization .NET | Aspose.Cells hide pivot row grand total

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The code loads an existing workbook, accesses the first pivot table, disables its row grand total by setting PivotTable.RowGrand to false, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook containing the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one pivot table on the worksheet
            if (sheet.PivotTables.Count > 0)
            {
                // Get the first pivot table
                PivotTable pivot = sheet.PivotTables[0];

                // Disable row grand totals
                pivot.RowGrand = false;
            }
            else
            {
                Console.WriteLine("No pivot tables found on the worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
