// Title: Create a pivot table with Aspose.Cells for .NET to group sales by region and calculate total sales
// AI Prompts: Write C# code using Aspose.Cells that builds a pivot table from a worksheet range, assigns "Region" as a row field, and sums the "Sales" column. | Show how to refresh and calculate the pivot table data after configuring fields with Aspose.Cells in a .NET project. | Demonstrate saving the workbook that contains the configured pivot table to an .xlsx file using Aspose.Cells.
// Common Searches: Aspose.Cells C# create pivot table that groups sales by region | How to add a row field and a summed data field to a pivot table with Aspose.Cells .NET | Refresh and calculate pivot table after field setup in Aspose.Cells C# | Save workbook with pivot table to Excel file using Aspose.Cells
// Tags: Aspose.Cells create pivot table from range | Aspose.Cells set row field Region | Aspose.Cells sum data field Sales | Aspose.Cells refresh calculate pivot data | Aspose.Cells save workbook as xlsx with pivot

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableExample
{
    // The example creates a new workbook, fills it with sample sales data, adds a separate worksheet for a pivot table, configures the 'Region' column as the row field and the 'Sales' column as a summed data field, refreshes and calculates the pivot, and finally saves the file as SalesByRegionPivot.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Prepare source data on the first worksheet
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Header row
            dataSheet.Cells["A1"].PutValue("Region");
            dataSheet.Cells["B1"].PutValue("Product");
            dataSheet.Cells["C1"].PutValue("Sales");

            // Sample data rows
            dataSheet.Cells["A2"].PutValue("North");
            dataSheet.Cells["B2"].PutValue("Widget");
            dataSheet.Cells["C2"].PutValue(1200);

            dataSheet.Cells["A3"].PutValue("South");
            dataSheet.Cells["B3"].PutValue("Gadget");
            dataSheet.Cells["C3"].PutValue(850);

            dataSheet.Cells["A4"].PutValue("North");
            dataSheet.Cells["B4"].PutValue("Gadget");
            dataSheet.Cells["C4"].PutValue(430);

            dataSheet.Cells["A5"].PutValue("East");
            dataSheet.Cells["B5"].PutValue("Widget");
            dataSheet.Cells["C5"].PutValue(670);

            dataSheet.Cells["A6"].PutValue("South");
            dataSheet.Cells["B6"].PutValue("Widget");
            dataSheet.Cells["C6"].PutValue(910);

            // -------------------------------------------------
            // 2. Add a new worksheet to host the pivot table
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // -------------------------------------------------
            // 3. Create the pivot table
            // -------------------------------------------------
            // Define the source range (including headers)
            string sourceRange = "=Data!A1:C6";

            // Destination cell for the pivot table (upper‑left corner)
            string destCell = "A3";

            // Add the pivot table and obtain its reference
            int pivotIndex = pivotSheet.PivotTables.Add(sourceRange, destCell, "SalesByRegion");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // -------------------------------------------------
            // 4. Configure fields:
            //    - Region as row field
            //    - Sales as data field (summed)
            // -------------------------------------------------
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Optional: display the pivot table in tabular form for readability
            pivotTable.ShowInTabularForm();

            // -------------------------------------------------
            // 5. Refresh and calculate the pivot data
            // -------------------------------------------------
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // -------------------------------------------------
            // 6. Save the workbook
            // -------------------------------------------------
            workbook.Save("SalesByRegionPivot.xlsx");
        }
    }
}
