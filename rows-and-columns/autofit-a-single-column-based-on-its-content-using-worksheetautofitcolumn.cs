using System;
using Aspose.Cells;

namespace AutoFitColumnDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data in column A (index 0)
            worksheet.Cells["A1"].PutValue("Short");
            worksheet.Cells["A2"].PutValue("This is a much longer piece of text that will require column width adjustment");
            worksheet.Cells["A3"].PutValue("Another long text entry to demonstrate AutoFitColumn");

            // Display column width before autofit
            Console.WriteLine($"Column width before AutoFitColumn: {worksheet.Cells.GetColumnWidth(0)}");

            // Auto‑fit column A (zero‑based index 0) using Worksheet.AutoFitColumn
            worksheet.AutoFitColumn(0);

            // Display column width after autofit
            Console.WriteLine($"Column width after AutoFitColumn: {worksheet.Cells.GetColumnWidth(0)}");

            // Save the workbook (save rule)
            workbook.Save("AutoFitColumnResult.xlsx");
        }
    }
}