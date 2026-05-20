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

            // -------------------------------------------------
            // Sample data that mimics the /Orders/Order element
            // -------------------------------------------------
            // Assume the XML mapping has been imported and the data resides in the first worksheet.
            // Columns: OrderID, Customer, Product, Quantity, Price
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Orders";

            // Header row
            dataSheet.Cells["A1"].PutValue("OrderID");
            dataSheet.Cells["B1"].PutValue("Customer");
            dataSheet.Cells["C1"].PutValue("Product");
            dataSheet.Cells["D1"].PutValue("Quantity");
            dataSheet.Cells["E1"].PutValue("Price");

            // Sample rows
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
            dataSheet.Cells["E4"].PutValue(500);

            dataSheet.Cells["A5"].PutValue(1004);
            dataSheet.Cells["B5"].PutValue("Charlie");
            dataSheet.Cells["C5"].PutValue("Laptop");
            dataSheet.Cells["D5"].PutValue(1);
            dataSheet.Cells["E5"].PutValue(1200);

            // -------------------------------------------------
            // Create a worksheet to host the pivot table
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotReport");

            // Define the source data range (including headers)
            string sourceData = "=Orders!A1:E5";

            // Add a new pivot table at cell A3 of the pivot sheet
            int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A3", "OrdersPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // -------------------------------------------------
            // Configure the pivot table fields
            // -------------------------------------------------
            // Row: Customer
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Customer");

            // Column: Product
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");

            // Data: Sum of Quantity
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Data: Sum of Price (optional, can be added as another data field)
            // pivotTable.AddFieldToArea(PivotFieldType.Data, "Price");

            // Optional: Show the pivot table in compact form for better readability
            pivotTable.ShowInCompactForm();

            // Refresh the pivot cache and calculate the results
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("OrdersPivotReport.xlsx");
        }
    }
}