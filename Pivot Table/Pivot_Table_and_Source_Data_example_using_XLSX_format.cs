using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Clothing");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Electronics");
            sheet.Cells["B4"].PutValue(200);
            sheet.Cells["A5"].PutValue("Food");
            sheet.Cells["B5"].PutValue(150);
            sheet.Cells["A6"].PutValue("Clothing");
            sheet.Cells["B6"].PutValue(70);

            // Add a pivot table using the source range A1:B6, place it at E3, and name it "SalesPivot"
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Category as row field, Amount as data field
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Display the current data source of the pivot table
            string[] source = pivot.GetSource();
            Console.WriteLine("Initial Pivot Table Source: " + source[0]);

            // Change the data source to a new range (A1:B4) – this demonstrates ChangeDataSource
            pivot.ChangeDataSource(new[] { "A1:B4" });

            // Verify the updated source
            source = pivot.GetSource();
            Console.WriteLine("Updated Pivot Table Source: " + source[0]);

            // Modify some source data to see the effect of refresh
            sheet.Cells["B2"].PutValue(130); // Food amount changed
            sheet.Cells["B3"].PutValue(90);  // Clothing amount changed

            // Refresh all pivot tables in the workbook to reflect the changes
            workbook.Worksheets.RefreshPivotTables();

            // Save the workbook to an XLSX file
            workbook.Save("PivotTableExample.xlsx");

            Console.WriteLine("Pivot table created and saved successfully.");
        }
    }
}