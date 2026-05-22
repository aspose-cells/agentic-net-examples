using System;
using Aspose.Cells;

namespace ColumnWidthDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate column A with short text
            cells["A1"].PutValue("Short");
            cells["A2"].PutValue("Data");

            // Populate column B with longer text that will require autofit
            cells["B1"].PutValue("This is a much longer piece of text that should cause the column to expand when AutoFitColumn is called");
            cells["B2"].PutValue("Another long text entry for demonstration purposes");

            // Manually set the width of column A (index 0) to 20 characters
            cells.SetColumnWidth(0, 20.0);

            // Auto‑fit column B (index 1) based on its content
            worksheet.AutoFitColumn(1);

            // Optional: display the resulting widths in the console
            Console.WriteLine($"Column A width (manual): {cells.GetColumnWidth(0)}");
            Console.WriteLine($"Column B width (auto‑fit): {cells.GetColumnWidth(1)}");

            // Save the workbook
            workbook.Save("ColumnWidthComparison.xlsx");
        }
    }
}