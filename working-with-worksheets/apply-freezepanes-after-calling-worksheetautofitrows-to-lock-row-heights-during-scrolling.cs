using System;
using System.IO;
using Aspose.Cells;

namespace FreezePanesAfterAutoFitRows
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data to demonstrate row height changes
            worksheet.Cells["A1"].PutValue("Short text");
            worksheet.Cells["A2"].PutValue("This is a longer piece of text that will cause the row height to increase when wrapped.");
            worksheet.Cells["A3"].PutValue("Another line with\nmultiple line breaks\nto test auto‑fit.");

            // Enable text wrapping for the cells so that row height can be affected
            Style wrapStyle = worksheet.Cells["A2"].GetStyle();
            wrapStyle.IsTextWrapped = true;
            worksheet.Cells["A2"].SetStyle(wrapStyle);
            worksheet.Cells["A3"].SetStyle(wrapStyle);

            // Auto‑fit all rows in the worksheet
            worksheet.AutoFitRows();

            // Freeze panes after auto‑fit to lock the row heights.
            // Freeze the first three rows (row index 3) and the first column (column index 0).
            // Parameters: row index, column index, number of frozen rows, number of frozen columns.
            worksheet.FreezePanes(3, 0, 3, 0);

            // Save the workbook to the desktop (adjust the path as needed)
            string outputPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "FreezePanesAfterAutoFitRows.xlsx");

            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}