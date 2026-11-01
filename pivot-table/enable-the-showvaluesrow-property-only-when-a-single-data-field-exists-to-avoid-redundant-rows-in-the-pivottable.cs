// Title: Enable ShowValuesRow only when a single data field exists in an Aspose.Cells PivotTable (C#)
// Common Searches: C# Aspose.Cells set ShowValuesRow only if pivot table has one data field | how to hide values row in Excel pivot table using Aspose.Cells when multiple data fields are present | conditional ShowValuesRow property in Aspose.Cells PivotTable example | Aspose.Cells pivot table DataFields count check before showing values row | remove redundant values row from a pivot table with Aspose.Cells .NET
// Tags: Aspose.Cells pivot table ShowValuesRow conditional | C# check PivotTable DataFields count Aspose | hide redundant values row Excel pivot Aspose.Cells | modify pivot table properties programmatically Aspose.Cells | Excel workbook pivot table manipulation using .NET | Aspose.Cells set ShowValuesRow based on data field count

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing Excel workbook, accesses the first worksheet and its first pivot table, and sets the ShowValuesRow property to true only when the pivot table contains exactly one data field, preventing redundant values rows before saving the workbook.
class PivotTableShowValuesRowExample
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of pivot tables in the worksheet
            PivotTableCollection pivotTables = worksheet.PivotTables;

            // Ensure there is at least one pivot table
            if (pivotTables != null && pivotTables.Count > 0)
            {
                // Work with the first pivot table (adjust index if needed)
                PivotTable pivotTable = pivotTables[0];

                // Show the values row only when a single data field exists
                pivotTable.ShowValuesRow = pivotTable.DataFields.Count == 1;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
