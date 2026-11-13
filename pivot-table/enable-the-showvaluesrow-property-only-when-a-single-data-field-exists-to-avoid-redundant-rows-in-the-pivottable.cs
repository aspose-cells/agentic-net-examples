// Title: Set ShowValuesRow for Aspose.Cells PivotTables only when a single data field exists (C#)
// Common Searches: C# Aspose.Cells set ShowValuesRow only if pivot table has one data field | How to hide the values row in an Excel pivot table using Aspose.Cells when multiple data fields are present | Aspose.Cells pivot table ShowValuesRow property conditional based on DataFields count | Programmatically enable values row for single-data-field pivot tables with Aspose.Cells .NET | Remove redundant values row in Excel pivot using Aspose.Cells C#
// Tags: Aspose.Cells ShowValuesRow conditional | C# pivot table single data field | Excel pivot hide redundant values row | set ShowValuesRow based on DataFields count | modify pivot layout with Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads a workbook, iterates through each PivotTable on the first worksheet, and sets the ShowValuesRow property to true only when the pivot contains exactly one data field, preventing redundant rows, then saves the workbook.
class PivotTableShowValuesRow
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume pivot tables are on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through all pivot tables in the worksheet
            foreach (PivotTable pivot in sheet.PivotTables)
            {
                // Show values row only when there is a single data field
                pivot.ShowValuesRow = pivot.DataFields.Count == 1;
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
