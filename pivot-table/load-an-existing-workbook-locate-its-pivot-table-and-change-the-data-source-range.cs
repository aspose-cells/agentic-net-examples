// Title: Change the data source range of a pivot table in an existing Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel file, locate its first pivot table, set the data source to a new range (e.g., C1:D10), refresh and save the workbook with Aspose.Cells in C#. | Generate C# code that uses Aspose.Cells to modify a pivot table’s source address, recalculate the data, and write the updated workbook to disk.
// Common Searches: asp.net aspose.cells change pivot table source address | c# update pivot table data range using Aspose.Cells | how to refresh a pivot table after changing its source with Aspose.Cells | Aspose.Cells ChangeDataSource method example for Excel pivot tables | programmatically set new data source for existing pivot table in .NET
// Tags: Aspose.Cells ChangeDataSource pivot table | C# update pivot table source range | Aspose.Cells refresh pivot table data | modify Excel pivot table data source .NET | Aspose.Cells recalculate pivot table after source change

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads InputWorkbook.xlsx, accesses the first worksheet's first pivot table, changes its data source to the range C1:D10 on the same sheet, refreshes and recalculates the pivot table, and saves the result as OutputWorkbook.xlsx.
class ChangePivotTableDataSource
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Assume the pivot table is in the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the collection of pivot tables in the worksheet
        PivotTableCollection pivotTables = worksheet.PivotTables;

        // Ensure there is at least one pivot table
        if (pivotTables.Count > 0)
        {
            // Access the first pivot table
            PivotTable pivotTable = pivotTables[0];

            // Define the new data source range (e.g., C1:D10 on the same sheet)
            // The array contains the range address and the sheet name
            string[] newDataSource = new string[] { "C1:D10", worksheet.Name };

            // Change the data source of the pivot table
            pivotTable.ChangeDataSource(newDataSource);

            // Refresh the pivot table to apply the new source and recalculate data
            pivotTable.RefreshData();
            pivotTable.CalculateData();
        }
        else
        {
            Console.WriteLine("No pivot tables found in the worksheet.");
        }

        // Save the modified workbook
        workbook.Save("OutputWorkbook.xlsx");
    }
}
