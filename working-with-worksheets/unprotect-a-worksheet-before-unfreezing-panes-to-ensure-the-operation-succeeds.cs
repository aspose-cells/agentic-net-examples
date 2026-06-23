using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Unprotect the worksheet (no password required)
        worksheet.Unprotect();

        // Unfreeze any frozen panes
        worksheet.UnFreezePanes();

        // Save the changes to a new file
        workbook.Save("output.xlsx");
    }
}