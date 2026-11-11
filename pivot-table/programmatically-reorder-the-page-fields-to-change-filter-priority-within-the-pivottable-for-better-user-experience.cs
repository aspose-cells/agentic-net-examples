// Title: Programmatically reorder PivotTable page fields to change filter priority using Aspose.Cells for .NET (C#)
// Common Searches: C# code to change the order of page fields in an Excel PivotTable with Aspose.Cells | How to set filter priority of PivotTable page fields using Aspose.Cells .NET | Move a PivotTable page field to a different position in a workbook via Aspose.Cells | Reorder Excel PivotTable page filters programmatically in C# | Aspose.Cells example for adjusting PivotTable page field sequence
// Tags: Aspose.Cells reorder pivot page fields | C# change pivot filter priority | Excel PivotTable page field manipulation .NET | programmatic pivot page field ordering | adjust pivot table filter sequence with Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an existing Excel workbook, accesses the first worksheet, retrieves a PivotTable named "PivotTable1", verifies that at least two page fields exist, moves the page field at index 1 to index 0 to change filter priority, and saves the modified workbook as a new file.
class PivotTablePageFieldReorder
{
    static void Main()
    {
        try
        {
            string inputPath = "InputWithPivot.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the PivotTable by its name
            string pivotName = "PivotTable1";
            PivotTable pivotTable = sheet.PivotTables[pivotName];
            if (pivotTable == null)
            {
                Console.WriteLine($"PivotTable '{pivotName}' not found.");
                return;
            }

            // Ensure there are at least two page fields to reorder
            if (pivotTable.PageFields.Count < 2)
            {
                Console.WriteLine("Not enough page fields to reorder.");
                return;
            }

            // Move the page field at index 1 to index 0
            PivotField fieldToMove = pivotTable.PageFields[1];
            pivotTable.PageFields.RemoveAt(1);
            pivotTable.PageFields.Insert(0, fieldToMove);

            // Save the updated workbook
            string outputPath = "OutputReorderedPivot.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
