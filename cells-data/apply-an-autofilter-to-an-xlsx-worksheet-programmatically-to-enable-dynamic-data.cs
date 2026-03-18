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

            // Populate sample data with a header row
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Category");
            worksheet.Cells["C1"].PutValue("Price");

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

            // Apply an AutoFilter to the range that includes the header and data rows
            worksheet.AutoFilter.Range = "A1:C5";

            // Example: filter the Category column (index 1) to show only "Electronics"
            worksheet.AutoFilter.Filter(1, "Electronics");

            // Refresh the filter to hide rows that do not meet the criteria
            worksheet.AutoFilter.Refresh();

            // Save the workbook with the AutoFilter applied
            workbook.Save("AutoFilterDemo.xlsx");
        }
    }
}