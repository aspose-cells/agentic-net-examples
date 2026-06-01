using System;
using Aspose.Cells;

namespace AutoFitRowExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add long text to cell A1 and enable text wrapping
            worksheet.Cells["A1"].PutValue("This is a very long piece of text that should cause the row height to increase when auto‑fitted.");
            Style style = worksheet.Cells["A1"].GetStyle();
            style.IsTextWrapped = true;
            worksheet.Cells["A1"].SetStyle(style);

            // Display row height before auto‑fit
            Console.WriteLine($"Row 0 height before AutoFitRow: {worksheet.Cells.Rows[0].Height}");

            // Auto‑fit the first row based on its content (operation rule)
            worksheet.AutoFitRow(0);

            // Display row height after auto‑fit
            Console.WriteLine($"Row 0 height after AutoFitRow: {worksheet.Cells.Rows[0].Height}");

            // Save the workbook (save rule)
            workbook.Save("AutoFitRowResult.xlsx");
        }
    }
}