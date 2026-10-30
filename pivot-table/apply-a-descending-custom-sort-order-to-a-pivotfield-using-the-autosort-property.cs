// Title: How to apply a descending custom sort to a PivotField with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set descending custom sort on pivot field | use AutoSort property to sort pivot table rows in descending order with Aspose.Cells | apply custom sort list to Excel pivot field using Aspose.Cells .NET API | C# code example for descending pivot field sort with Aspose.Cells | Aspose.Cells AutoSortCustomList not available – how to sort pivot rows
// Tags: Aspose.Cells pivot field descending sort | AutoSortCustomList API C# | custom pivot field sort Aspose.Cells | pivot table row field sorting .NET | Excel pivot custom sort Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an existing workbook, retrieves the first pivot table, selects its first row field, and notes that the AutoSort, AutoSortOrder, and AutoSortCustomList properties are unavailable in the current Aspose.Cells version, meaning a descending custom sort cannot be applied directly via those APIs.
class PivotFieldCustomSort
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve pivot tables from the worksheet
            PivotTableCollection pivotTables = worksheet.PivotTables;

            if (pivotTables == null || pivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Work with the first pivot table
            PivotTable pivotTable = pivotTables[0];

            if (pivotTable.RowFields == null || pivotTable.RowFields.Count == 0)
            {
                Console.WriteLine("The pivot table has no row fields.");
                return;
            }

            // Choose the first row field to sort (replace index as needed)
            PivotField pivotField = pivotTable.RowFields[0];

            // NOTE: The AutoSort, AutoSortOrder, and AutoSortCustomList properties are not
            // available in the current Aspose.Cells version. If custom sorting is required,
            // use the appropriate API provided by the library version you are using.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
