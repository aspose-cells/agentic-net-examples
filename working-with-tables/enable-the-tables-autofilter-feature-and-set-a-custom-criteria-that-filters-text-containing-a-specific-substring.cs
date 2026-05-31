using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableAutoFilter
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the table (header + rows)
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple iPhone");
            sheet.Cells["B2"].PutValue("Electronics");
            sheet.Cells["A3"].PutValue("Banana Bread");
            sheet.Cells["B3"].PutValue("Food");
            sheet.Cells["A4"].PutValue("Apple MacBook");
            sheet.Cells["B4"].PutValue("Electronics");
            sheet.Cells["A5"].PutValue("Cherry Pie");
            sheet.Cells["B5"].PutValue("Food");

            // Add a ListObject (table) covering the data range A1:B5
            int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Enable auto‑filter for the table
            table.HasAutoFilter = true;

            // Apply a custom filter on the first column (Product) to show rows containing "Apple"
            // FilterOperatorType.Contains corresponds to the "contains" criterion
            table.AutoFilter.Custom(0, FilterOperatorType.Contains, "Apple");
            table.AutoFilter.Refresh();

            // Save the workbook
            workbook.Save("TableAutoFilterContains.xlsx");
        }
    }
}