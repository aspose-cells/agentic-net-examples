using System;
using Aspose.Cells;

namespace AsposeCellsTop10FilterDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: Product names and their sales figures
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");

            string[] products = { "Apple", "Banana", "Orange", "Grape", "Mango", "Peach", "Cherry", "Pear", "Plum", "Kiwi", "Lemon", "Lime", "Pineapple", "Watermelon", "Strawberry" };
            int[] sales =      { 120,   85,      150,    60,    200,   95,     70,      55,    65,    40,    30,    25,      180,        210,          110 };

            for (int i = 0; i < products.Length; i++)
            {
                sheet.Cells[i + 2, 0].PutValue(products[i]); // Column A
                sheet.Cells[i + 2, 1].PutValue(sales[i]);   // Column B
            }

            // Define the autofilter range (including header row)
            int lastRow = products.Length + 1; // +1 for header
            sheet.AutoFilter.Range = $"A1:B{lastRow}";

            // Apply a Top 10 filter on the Sales column (field index 1)
            // isTop = true (show highest), isPercent = false (use count), itemCount = 10
            sheet.AutoFilter.FilterTop10(fieldIndex: 1, isTop: true, isPercent: false, itemCount: 10);

            // Refresh the filter to hide rows that do not meet the criteria
            sheet.AutoFilter.Refresh();

            // Save the workbook
            workbook.Save("Top10Products.xlsx");
        }
    }
}