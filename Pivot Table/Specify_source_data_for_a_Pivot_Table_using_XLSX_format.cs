using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotSourceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (XLSX by default)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Beverage");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Food");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["A5"].PutValue("Beverage");
            sheet.Cells["B5"].PutValue(70);

            // Add a pivot table using the initial source range A1:B5
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure fields (Category as row, Amount as data)
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Display the current data source
            string[] currentSource = pivot.DataSource;
            Console.WriteLine("Current Pivot Source: " + currentSource[0]);

            // Change the data source to a new range C1:D5 (add extra columns for demonstration)
            // First, add the new data range
            sheet.Cells["C1"].PutValue("Region");
            sheet.Cells["D1"].PutValue("Quantity");
            sheet.Cells["C2"].PutValue("North");
            sheet.Cells["D2"].PutValue(30);
            sheet.Cells["C3"].PutValue("South");
            sheet.Cells["D3"].PutValue(20);
            sheet.Cells["C4"].PutValue("North");
            sheet.Cells["D4"].PutValue(45);
            sheet.Cells["C5"].PutValue("South");
            sheet.Cells["D5"].PutValue(25);

            // Use ChangeDataSource to point the pivot table to the new range
            // The source array can contain the range and optionally the sheet name
            pivot.ChangeDataSource(new string[] { "C1:D5", sheet.Name });

            // Refresh the pivot table to apply the new source
            pivot.RefreshData();
            pivot.CalculateData();

            // Verify the updated source
            string[] updatedSource = pivot.DataSource;
            Console.WriteLine("Updated Pivot Source: " + updatedSource[0]);

            // Save the workbook in XLSX format
            workbook.Save("PivotTableSourceDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}