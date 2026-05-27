using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set the left margin to 0.5 inches
        sheet.PageSetup.LeftMarginInch = 0.5;

        // Add sample data (optional, just to have content in the sheet)
        sheet.Cells["A1"].PutValue("Left margin set to 0.5 inches");

        // Create save options for Excel 97‑2003 format
        XlsSaveOptions saveOptions = new XlsSaveOptions();

        // Save the workbook as an .xls file using the specified options
        workbook.Save("output.xls", saveOptions);
    }
}