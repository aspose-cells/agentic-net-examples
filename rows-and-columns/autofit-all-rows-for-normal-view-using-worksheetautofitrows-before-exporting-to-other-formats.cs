using System;
using Aspose.Cells;

class AutoFitRowsExample
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data that may require row height adjustment
        sheet.Cells["A1"].PutValue("This is a long text that will cause the row to expand when AutoFitRows is applied.");
        sheet.Cells["A2"].PutValue("Another line with\nmultiple line breaks\nto test row height auto-fitting.");
        sheet.Cells["B1"].PutValue("Short text");
        sheet.Cells["B2"].PutValue("Another short text");

        // Enable text wrapping for the cell containing line breaks
        Style wrapStyle = sheet.Cells["A2"].GetStyle();
        wrapStyle.IsTextWrapped = true;
        sheet.Cells["A2"].SetStyle(wrapStyle);

        // Auto‑fit all rows in the worksheet before exporting (feature rule)
        sheet.AutoFitRows();

        // Save the workbook to a PDF file (lifecycle rule: save)
        workbook.Save("AutoFitRowsOutput.pdf", SaveFormat.Pdf);
    }
}