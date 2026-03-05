using System;
using Aspose.Cells;

namespace AsposeCellsAutofilterDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (XLSX format by default)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Category");
            sheet.Cells["C1"].PutValue("Price");

            // Add sample data rows
            sheet.Cells["A2"].PutValue("Laptop");
            sheet.Cells["B2"].PutValue("Electronics");
            sheet.Cells["C2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("Shirt");
            sheet.Cells["B3"].PutValue("Clothing");
            sheet.Cells["C3"].PutValue(35);

            sheet.Cells["A4"].PutValue("Phone");
            sheet.Cells["B4"].PutValue("Electronics");
            sheet.Cells["C4"].PutValue(800);

            sheet.Cells["A5"].PutValue("Jeans");
            sheet.Cells["B5"].PutValue("Clothing");
            sheet.Cells["C5"].PutValue(60);

            // Apply AutoFilter to the range that includes the header and data rows
            sheet.AutoFilter.Range = "A1:C5";

            // Filter the "Category" column (index 1) to show only "Electronics"
            sheet.AutoFilter.Filter(1, "Electronics");

            // Refresh the filter to hide rows that do not meet the criteria
            sheet.AutoFilter.Refresh();

            // Save the workbook as an XLSX file
            workbook.Save("AutofilterDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}