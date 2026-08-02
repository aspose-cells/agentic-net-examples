using System;
using Aspose.Cells;

namespace ColumnWidthDemo
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate column B (index 1) with varying length text
            cells["B1"].PutValue("Short");
            cells["B2"].PutValue("A considerably longer piece of text");
            cells["B3"].PutValue("Medium length");

            // Set a fixed width for column A (index 0)
            // Width is specified in characters; 20.0 is an arbitrary example
            cells.SetColumnWidth(0, 20.0);

            // Auto‑fit column B (index 1) based on its content
            worksheet.AutoFitColumn(1);

            // Optional: display widths before saving (for debugging)
            Console.WriteLine($"Column A width (fixed): {cells.GetColumnWidth(0)}");
            Console.WriteLine($"Column B width (auto‑fitted): {cells.GetColumnWidth(1)}");

            // Save the workbook
            workbook.Save("ColumnWidthComparison.xlsx");
        }
    }
}