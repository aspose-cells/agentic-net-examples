using System;
using Aspose.Cells;

class FilterDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate headers
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Category");
        worksheet.Cells["C1"].PutValue("Price");

        // Populate sample data
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue("Fruits");
        worksheet.Cells["C2"].PutValue(2.5);

        worksheet.Cells["A3"].PutValue("Carrot");
        worksheet.Cells["B3"].PutValue("Vegetables");
        worksheet.Cells["C3"].PutValue(1.2);

        worksheet.Cells["A4"].PutValue("Banana");
        worksheet.Cells["B4"].PutValue("Fruits");
        worksheet.Cells["C4"].PutValue(1.8);

        worksheet.Cells["A5"].PutValue("Steak");
        worksheet.Cells["B5"].PutValue("Meat");
        worksheet.Cells["C5"].PutValue(5.0);

        // Define the auto‑filter range (including header row)
        worksheet.AutoFilter.Range = "A1:C5";

        // Apply a custom filter: show only rows where Price (column index 2) > 2.0
        worksheet.AutoFilter.Custom(2, FilterOperatorType.GreaterThan, 2.0);

        // Refresh the filter to hide rows that do not meet the criteria
        worksheet.AutoFilter.Refresh();

        // Save the filtered workbook
        workbook.Save("FilteredProducts.xlsx");
    }
}