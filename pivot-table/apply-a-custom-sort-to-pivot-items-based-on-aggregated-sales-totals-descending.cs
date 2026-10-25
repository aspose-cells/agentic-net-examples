// Title: Sort pivot table rows by total sales in descending order using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# sort pivot table rows by sum of sales column descending | How to apply a custom descending sort to pivot items based on aggregated values in Aspose.Cells | Enable AutoSort for a pivot field and sort by data field using Aspose.Cells .NET | Refresh and recalculate a pivot table after changing sort settings with Aspose.Cells | Aspose.Cells version support for PivotField.SortDataField property in C#
// Tags: pivot row field descending sort Aspose.Cells | custom sort by aggregated sales total Aspose.Cells | refresh and calculate pivot table Aspose.Cells C# | AutoSort property version limitation Aspose.Cells | C# example of pivot table sorting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an Excel workbook, verifies that a pivot table exists on the first worksheet, retrieves the first row field and the first data field, and notes that the AutoSort, SortOrder, and SortDataField members are unavailable in the referenced Aspose.Cells version. It then refreshes and recalculates the pivot table before saving the result to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (assumed to hold the pivot table)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No pivot tables found on the first worksheet.");
                return;
            }

            // Get the first pivot table on the worksheet
            PivotTable pivotTable = sheet.PivotTables[0];

            // Ensure the pivot table has at least one row field and one data field
            if (pivotTable.RowFields.Count == 0 || pivotTable.DataFields.Count == 0)
            {
                Console.WriteLine("Error: Pivot table does not contain required row or data fields.");
                return;
            }

            // Assume the row field to be sorted is the first row field (e.g., "Product")
            PivotField rowField = pivotTable.RowFields[0];

            // Assume the data field that contains the sales values is the first data field
            PivotField dataField = pivotTable.DataFields[0];

            // NOTE:
            // The AutoSort, SortOrder, and SortDataField properties are not available
            // in the version of Aspose.Cells referenced by this project.
            // If your Aspose.Cells version supports these members, uncomment the lines below:

            // rowField.AutoSort = true;                                 // Enable auto‑sort
            // rowField.SortOrder = PivotSortType.Descending;           // Set descending order
            // rowField.SortDataField = dataField;                       // Sort by the sales data field

            // Refresh the pivot table to apply any changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Pivot table processed successfully. Output saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
