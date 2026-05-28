using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Prepare source data (Region, Sales)
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Header
            dataSheet.Cells["A1"].PutValue("Region");
            dataSheet.Cells["B1"].PutValue("Sales");

            // Sample rows
            dataSheet.Cells["A2"].PutValue("North");
            dataSheet.Cells["B2"].PutValue(1200);
            dataSheet.Cells["A3"].PutValue("South");
            dataSheet.Cells["B3"].PutValue(850);
            dataSheet.Cells["A4"].PutValue("East");
            dataSheet.Cells["B4"].PutValue(950);
            dataSheet.Cells["A5"].PutValue("West");
            dataSheet.Cells["B5"].PutValue(1100);
            dataSheet.Cells["A6"].PutValue("North");
            dataSheet.Cells["B6"].PutValue(700);
            dataSheet.Cells["A7"].PutValue("South");
            dataSheet.Cells["B7"].PutValue(650);
            dataSheet.Cells["A8"].PutValue("East");
            dataSheet.Cells["B8"].PutValue(400);
            dataSheet.Cells["A9"].PutValue("West");
            dataSheet.Cells["B9"].PutValue(500);

            // -------------------------------------------------
            // 2. Add a worksheet for the pivot table
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // -------------------------------------------------
            // 3. Create the pivot table
            // -------------------------------------------------
            // Source range: A1:B9 on the Data sheet
            // Destination cell: A3 on the pivot sheet
            int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:B9", "A3", "SalesByRegion");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // -------------------------------------------------
            // 4. Configure fields:
            //    - Row field: Region
            //    - Data field: Sales (summed)
            // -------------------------------------------------
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Ensure the data field uses Sum aggregation (default is Sum, but set explicitly)
            if (pivotTable.DataFields.Count > 0)
            {
                pivotTable.DataFields[0].Function = ConsolidationFunction.Sum;
                pivotTable.DataFields[0].NumberFormat = "$#,##0";
            }

            // Optional: display the pivot in tabular form for readability
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