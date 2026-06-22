using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data that requires row height adjustment
        worksheet.Cells["A1"].PutValue("This is a long text that should cause the row height to increase automatically when auto‑fit is applied.");
        worksheet.Cells["A2"].PutValue("Short text");

        // Enable text wrapping so the long text occupies multiple lines
        Style wrapStyle = worksheet.Cells["A1"].GetStyle();
        wrapStyle.IsTextWrapped = true;
        worksheet.Cells["A1"].SetStyle(wrapStyle);

        // Auto‑fit all rows in the worksheet; this makes the row heights adapt to the cell contents
        worksheet.AutoFitRows();

        // Save the workbook to verify the result
        workbook.Save("AutoFitRowsResult.xlsx");
    }
}