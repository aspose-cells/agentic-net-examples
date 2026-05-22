using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add sample data to column A (index 0) with varying lengths
        cells["A1"].PutValue("Short");
        cells["A2"].PutValue("This is a longer piece of text that will cause the column to expand");
        cells["A3"].PutValue("Medium length");

        // Auto‑fit column A for rows 0 to 2 (zero‑based indices)
        worksheet.AutoFitColumn(0, 0, 2);

        // Fine‑tune the column width to an exact pixel value (e.g., 150 pixels)
        cells.SetColumnWidthPixel(0, 150);

        // Save the workbook to a file
        workbook.Save("AutoFitAndFineTuneColumn.xlsx");
    }
}