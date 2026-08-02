using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a worksheet that will hold the source data
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Orders";

            // Populate sample data that mimics the /Orders/Order XML element mapping
            // Columns: OrderID, Customer, Amount
            dataSheet.Cells["A1"].PutValue("OrderID");
            dataSheet.Cells["B1"].PutValue("Customer");
            dataSheet.Cells["C1"].PutValue("Amount");

            // Sample rows
            dataSheet.Cells["A2"].PutValue(1001);
            dataSheet.Cells["B2"].PutValue("Alice");
            dataSheet.Cells["C2"].PutValue(250);

            dataSheet.Cells["A3"].PutValue(1002);
            dataSheet.Cells["B3"].PutValue("Bob");
            dataSheet.Cells["C3"].PutValue(150);

            dataSheet.Cells["A4"].PutValue(1003);
            dataSheet.Cells["B4"].PutValue("Alice");
            dataSheet.Cells["C4"].PutValue(300);

            dataSheet.Cells["A5"].PutValue(1004);
            dataSheet.Cells["B5"].PutValue("Charlie");
            dataSheet.Cells["C5"].PutValue(400);

            // Add a new worksheet that will contain the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Define the source data range for the pivot table
            // Using the full range including headers
            string sourceData = "=Orders!A1:C5";

            // Add the pivot table to the pivot sheet at cell A3
            int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A3", "OrdersSummary");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot table:
            // - Row field: Customer
            // - Data field: Amount (sum)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Customer");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Optional: set the pivot table to display in compact form
            pivotTable.ShowInCompactForm();

            // Refresh the pivot cache and calculate the data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("OrdersPivotTable.xlsx");
        }
    }
}