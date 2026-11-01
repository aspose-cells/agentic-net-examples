// Title: How to set PivotTable MissingItemsLimit to Max using Aspose.Cells in C# (.NET)
// Common Searches: Aspose.Cells C# set PivotTable MissingItemsLimit to include all items | C# example using PivotMissingItemLimitType.Max with Aspose.Cells | Refresh an Aspose.Cells pivot table so it shows every possible item | Configure missing items limit for a pivot table in Aspose.Cells .NET
// Tags: Aspose.Cells PivotTable MissingItemsLimit Max | C# PivotMissingItemLimitType usage | Aspose.Cells refresh pivot table all items | set missing items limit Aspose.Cells | pivot table configuration Aspose.Cells .NET

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;

// The sample creates a workbook, adds sample data, inserts a PivotTable, sets its MissingItemsLimit property to PivotMissingItemLimitType.Max to ensure all possible items appear after a refresh, refreshes and calculates the pivot data, and saves the file as PivotTableMissingItemsLimit.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(100);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(200);
            dataSheet.Cells["A4"].PutValue("C");
            dataSheet.Cells["B4"].PutValue(300);
            // Additional rows can be added as needed

            // Add a new worksheet for the pivot table
            int pivotSheetIndex = workbook.Worksheets.Add();
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
            pivotSheet.Name = "Pivot";

            // Define the source data range for the pivot table
            string sourceData = "Data!A1:B4";

            // Add the pivot table to the pivot sheet
            int pivotTableIndex = pivotSheet.PivotTables.Add(sourceData, "C1", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotTableIndex];

            // Set MissingItemsLimit to include all possible items during refresh
            pivotTable.MissingItemsLimit = PivotMissingItemLimitType.Max;

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTableMissingItemsLimit.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
