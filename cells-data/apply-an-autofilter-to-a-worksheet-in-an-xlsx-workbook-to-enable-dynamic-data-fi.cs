using System;
using Aspose.Cells;

namespace AutoFilterDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add header row
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Category");
            worksheet.Cells["C1"].PutValue("Price");

            // Add sample data
            worksheet.Cells["A2"].PutValue("Laptop");
            worksheet.Cells["B2"].PutValue("Electronics");
            worksheet.Cells["C2"].PutValue(1200);

            worksheet.Cells["A3"].PutValue("Shirt");
            worksheet.Cells["B3"].PutValue("Clothing");
            worksheet.Cells["C3"].PutValue(45);

            worksheet.Cells["A4"].PutValue("Phone");
            worksheet.Cells["B4"].PutValue("Electronics");
            worksheet.Cells["C4"].PutValue(800);

            worksheet.Cells["A5"].PutValue("Pants");
            worksheet.Cells["B5"].PutValue("Clothing");
            worksheet.Cells["C5"].PutValue(60);

            // Apply AutoFilter to the range that includes headers and data
            worksheet.AutoFilter.Range = "A1:C5";

            // Example filter: show only rows where Category = "Electronics"
            worksheet.AutoFilter.Filter(1, "Electronics");
            worksheet.AutoFilter.Refresh();

            // Save the workbook with the applied AutoFilter
            workbook.Save("AutoFilterDemo.xlsx");
        }
    }
}