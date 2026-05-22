using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add long wrapped text to demonstrate row auto‑fit
        worksheet.Cells["A1"].PutValue("This is a very long text that will normally cause the row to become excessively tall after auto‑fit.");
        Style style = worksheet.Cells["A1"].GetStyle();
        style.IsTextWrapped = true;
        worksheet.Cells["A1"].SetStyle(style);

        // Set an initial small row height (optional)
        worksheet.Cells.SetRowHeight(0, 10);

        // Define AutoFitterOptions with a maximum row height limit
        AutoFitterOptions options = new AutoFitterOptions
        {
            MaxRowHeight = 40, // limit row height to 40 points
            OnlyAuto = true    // apply only to rows without custom height
        };

        // Auto‑fit rows using the specified options
        worksheet.AutoFitRows(options);

        // Save the workbook
        workbook.Save("MaxRowHeightDemo.xlsx");
    }
}