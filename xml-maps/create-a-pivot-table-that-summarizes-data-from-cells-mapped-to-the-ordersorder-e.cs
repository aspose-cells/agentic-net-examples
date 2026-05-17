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

            // Get the first worksheet (will hold the source data)
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Orders";

            // ------------------------------------------------------------
            // Sample data that mimics an XML mapping to /Orders/Order
            // Columns: OrderID, Customer, Product, Quantity, Price
            // ------------------------------------------------------------
            dataSheet.Cells["A1"].PutValue("OrderID");
            dataSheet.Cells["B1"].PutValue("Customer");
            dataSheet.Cells["C1"].PutValue("Product");
            dataSheet.Cells["D1"].PutValue("Quantity");
            dataSheet.Cells["E1"].PutValue("Price");

            // Populate a few rows
            dataSheet.Cells["A2"].PutValue(1001);
            dataSheet.Cells["B2"].PutValue("Alice");
            dataSheet.Cells["C2"].PutValue("Laptop");
            dataSheet.Cells["D2"].PutValue(2);
            dataSheet.Cells["E2"].PutValue(1200);

            dataSheet.Cells["A3"].PutValue(1002);
            dataSheet.Cells["B3"].PutValue("Bob");
            dataSheet.Cells["C3"].PutValue("Phone");
            dataSheet.Cells["D3"].PutValue(5);
            dataSheet.Cells["E3"].PutValue(800);

            dataSheet.Cells["A4"].PutValue(1003);
            dataSheet.Cells["B4"].PutValue("Alice");
            dataSheet.Cells["C4"].PutValue("Tablet");
            dataSheet.Cells["D4"].PutValue(3);
            dataSheet.Cells["E4"].PutValue(400);

            dataSheet.Cells["A5"].PutValue(1004);
            dataSheet.Cells["B5"].PutValue("Charlie");
            dataSheet.Cells["C5"].PutValue("Laptop");
            dataSheet.Cells["D5"].PutValue(1);
            dataSheet.Cells["E5"].PutValue(1200);

            // ------------------------------------------------------------
            // Add a new worksheet that will contain the PivotTable
            // ------------------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("OrdersPivot");

            // Define the source data range for the pivot table
            // Using the full range including headers (A1:E5)
            string sourceData = "=Orders!A1:E5";

            // Destination cell where the pivot table will start
            string destCell = "A3";

            // Name of the pivot table
            string pivotName = "OrdersSummary";

            // Add the pivot table to the pivot sheet
            int pivotIndex = pivotSheet.PivotTables.Add(sourceData, destCell, pivotName);
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // ------------------------------------------------------------
            // Configure the pivot table fields
            // Row: Customer
            // Column: Product
            // Data: Sum of (Quantity * Price) -> we will add Quantity and Price separately,
            // then use a calculated field for TotalAmount
            // ------------------------------------------------------------
            // Add row and column fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Customer");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");

            // Add Quantity and Price as data fields (they will be summed)
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Price");

            // Create a calculated field: TotalAmount = Quantity * Price
            // The formula syntax for calculated fields follows Excel's format.
            // Use field names enclosed in brackets.
            string formula = "Quantity*Price";
            pivotTable.AddCalculatedField("TotalAmount", formula, true); // true => drag to data area

            // Optionally, remove the individual Quantity and Price data fields
            // to keep only the calculated total in the data area.
            pivotTable.RemoveField(PivotFieldType.Data, "Quantity");
            pivotTable.RemoveField(PivotFieldType.Data, "Price");

            // Refresh the pivot cache and calculate the data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("OrdersPivotTable.xlsx");
        }
    }
}