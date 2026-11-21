// Title: How to programmatically reorder row PivotFields in an Aspose.Cells PivotTable using C#
// Common Searches: C# code to change the order of row fields in an Aspose.Cells pivot table | move pivot table row field index with Aspose.Cells .NET API | reorder PivotTable row hierarchy programmatically in Excel using Aspose.Cells | Aspose.Cells example for shifting row PivotFields positions | adjust row area field sequence in a pivot table with C#
// Tags: Aspose.Cells pivot table row field ordering | C# move PivotTable RowFields method | programmatic pivot hierarchy reordering Aspose.Cells | Excel pivot layout modification .NET | change pivot row hierarchy using Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// Loads an existing workbook, checks for a pivot table, and if the table contains more than two row fields, moves the field at index 0 to index 2 to alter the row hierarchy order before saving the updated file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table (adjust index if needed)
            PivotTable pivotTable = sheet.PivotTables[0];

            // Reorder row fields if there are enough fields
            if (pivotTable.RowFields.Count > 2)
            {
                // Move field from index 0 to index 2
                pivotTable.RowFields.Move(0, 2);
            }

            // Save the workbook with the updated pivot table layout
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
