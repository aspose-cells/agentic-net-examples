using System;
using Aspose.Cells;

class AutoFitRowsBeforeFreezeDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate cells with data that will affect row heights
        worksheet.Cells["A1"].PutValue("This is a long text that will cause the row height to increase when wrapped.");
        worksheet.Cells["A2"].PutValue("Short text");
        worksheet.Cells["A3"].PutValue("Another long text\nwith line breaks to demonstrate auto‑fit rows.");

        // Enable text wrapping for the cells that contain long text
        Style wrapStyle = worksheet.Cells["A1"].GetStyle();
        wrapStyle.IsTextWrapped = true;
        worksheet.Cells["A1"].SetStyle(wrapStyle);
        worksheet.Cells["A3"].SetStyle(wrapStyle);

        // Auto‑fit all rows to adjust heights based on the wrapped content
        worksheet.AutoFitRows();

        // Freeze panes at cell C3 (row index 2, column index 2) with 3 frozen rows and 3 frozen columns
        // This keeps the top rows and left columns visible while scrolling
        worksheet.FreezePanes("C3", 3, 3);

        // Save the workbook
        workbook.Save("AutoFitRowsAndFreeze.xlsx");
    }
}