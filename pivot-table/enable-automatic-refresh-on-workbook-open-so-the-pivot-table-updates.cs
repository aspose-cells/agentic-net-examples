using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotAutoRefresh
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Sales");
            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["B2"].PutValue(1000);
            dataSheet.Cells["A3"].PutValue("Banana");
            dataSheet.Cells["B3"].PutValue(2000);
            dataSheet.Cells["A4"].PutValue("Orange");
            dataSheet.Cells["B4"].PutValue(3000);

            // Add a pivot table based on the data range
            int pivotIndex = dataSheet.PivotTables.Add("A1:B4", "E3", "SalesPivot");
            PivotTable pivotTable = dataSheet.PivotTables[pivotIndex];

            // Configure the pivot table (Product as rows, Sales as data)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

            // Enable automatic refresh when the workbook is opened
            pivotTable.RefreshDataOnOpeningFile = true;

            // Save the workbook
            workbook.Save("PivotTable_AutoRefresh.xlsx");
        }
    }
}