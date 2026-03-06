using System;
using Aspose.Cells;

namespace AsposeCellsAutofilterDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Category");
            sheet.Cells["C1"].PutValue("Price");

            sheet.Cells["A2"].PutValue("Laptop");
            sheet.Cells["B2"].PutValue("Electronics");
            sheet.Cells["C2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("Shirt");
            sheet.Cells["B3"].PutValue("Clothing");
            sheet.Cells["C3"].PutValue(45);

            sheet.Cells["A4"].PutValue("Phone");
            sheet.Cells["B4"].PutValue("Electronics");
            sheet.Cells["C4"].PutValue(800);

            sheet.Cells["A5"].PutValue("Jeans");
            sheet.Cells["B5"].PutValue("Clothing");
            sheet.Cells["C5"].PutValue(60);

            // Apply an AutoFilter to the header range (A1:C5)
            sheet.AutoFilter.Range = "A1:C5";

            // Filter the "Category" column (field index 1) for "Electronics"
            sheet.AutoFilter.Filter(1, "Electronics");

            // Refresh the filter to hide non‑matching rows
            sheet.AutoFilter.Refresh();

            // Save the workbook in XLSX format (lifecycle: save)
            workbook.Save("FilteredProducts.xlsx", SaveFormat.Xlsx);
        }
    }
}