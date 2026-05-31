using System;
using Aspose.Cells;

namespace AutoFitColumnExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data to column A (index 0)
            worksheet.Cells["A1"].PutValue("Short");
            worksheet.Cells["A2"].PutValue("This is a much longer text that should cause the column to widen automatically");
            worksheet.Cells["A3"].PutValue("Medium length text");
            worksheet.Cells["A4"].PutValue("Another very long piece of text to demonstrate auto‑fit functionality");

            // Auto‑fit column A (zero‑based index 0)
            worksheet.AutoFitColumn(0);

            // Save the workbook
            workbook.Save("AutoFitColumnDemo.xlsx");
        }
    }
}