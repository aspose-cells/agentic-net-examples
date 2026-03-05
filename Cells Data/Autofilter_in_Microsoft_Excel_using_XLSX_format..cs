using System;
using Aspose.Cells;

namespace AsposeCellsAutofilterDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data with a header row
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Category");
            cells["C1"].PutValue("Price");

            cells["A2"].PutValue("Laptop");
            cells["B2"].PutValue("Electronics");
            cells["C2"].PutValue(1200);

            cells["A3"].PutValue("Shirt");
            cells["B3"].PutValue("Clothing");
            cells["C3"].PutValue(45);

            cells["A4"].PutValue("Phone");
            cells["B4"].PutValue("Electronics");
            cells["C4"].PutValue(800);

            cells["A5"].PutValue("Pants");
            cells["B5"].PutValue("Clothing");
            cells["C5"].PutValue(60);

            // Apply an AutoFilter to the header range (A1:C5)
            worksheet.AutoFilter.Range = "A1:C5";

            // Filter the 'Category' column (index 1) to show only 'Electronics'
            worksheet.AutoFilter.Filter(1, "Electronics");

            // Refresh the filter to hide non‑matching rows
            worksheet.AutoFilter.Refresh();

            // Save the workbook in XLSX format
            workbook.Save("AutofilterDemo.xlsx");
        }
    }
}