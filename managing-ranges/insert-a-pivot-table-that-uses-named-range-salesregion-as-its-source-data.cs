// Title: Create a Pivot Table from the Named Range "SalesRegion" Using Aspose.Cells in C#
// Description: Shows how to define a named range called SalesRegion for cells A1:B5, add a second worksheet, insert a pivot table that uses this named range as its source, configure Region as a row field and Sales as a data field, refresh the cache, calculate the results, and save the workbook.
// Keywords: Aspose.Cells pivot table example | C# named range pivot | SalesRegion named range | create pivot table Aspose.Cells | use named range as pivot source | Aspose.Cells workbook tutorial | pivot cache refresh Aspose | calculate pivot data C#
// Common Searches: Aspose.Cells create pivot table from named range C# | how to use named range as pivot source Aspose.Cells | C# example pivot table SalesRegion Aspose | Aspose.Cells add named range and pivot table | pivot table source range named range Aspose
// Developer Intent: Generate a pivot table whose source data is the named range "SalesRegion".
// Use Cases: Define a reusable named range for sales data and build a pivot report on a separate sheet. | Automatically update the pivot when new rows are added to the SalesRegion range. | Separate raw data entry from analytical summaries within the same workbook.
// AI Prompts: Provide C# code that creates a named range "SalesRegion" for A1:B5 and builds a pivot table from it using Aspose.Cells. | Explain the steps to replace a direct range reference with a named range when adding a pivot table in Aspose.Cells. | Show how to refresh and calculate a pivot table after modifying its named‑range source in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    // Shows how to define a named range called SalesRegion for cells A1:B5, add a second worksheet, insert a pivot table that uses this named range as its source, configure Region as a row field and Sales as a data field, refresh the cache, calculate the results, and save the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Prepare source data on the first worksheet
                // -------------------------------------------------
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "SourceData";

                // Sample data (Region, Sales)
                dataSheet.Cells["A1"].PutValue("Region");
                dataSheet.Cells["B1"].PutValue("Sales");
                dataSheet.Cells["A2"].PutValue("North");
                dataSheet.Cells["B2"].PutValue(1200);
                dataSheet.Cells["A3"].PutValue("South");
                dataSheet.Cells["B3"].PutValue(850);
                dataSheet.Cells["A4"].PutValue("East");
                dataSheet.Cells["B4"].PutValue(970);
                dataSheet.Cells["A5"].PutValue("West");
                dataSheet.Cells["B5"].PutValue(1100);

                // -------------------------------------------------
                // Add a new worksheet that will contain the pivot table
                // -------------------------------------------------
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Add the pivot table using a direct range reference as source data.
                // Parameters: sourceData, destination cell (upper‑left corner), table name
                int pivotIndex = pivotSheet.PivotTables.Add("'SourceData'!A1:B5", "A1", "SalesRegionPivot");

                // Retrieve the created pivot table
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure the pivot table (Region as row, Sales as data)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh the pivot cache and calculate the pivot data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                workbook.Save("PivotTableOutput.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
