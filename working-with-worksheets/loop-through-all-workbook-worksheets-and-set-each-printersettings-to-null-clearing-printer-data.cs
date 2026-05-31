using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (adjust the path as needed)
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Clear printer data by setting the PrinterSettings property to null
            sheet.PageSetup.PrinterSettings = null;
        }

        // Save the workbook after clearing printer settings
        workbook.Save("output.xlsx");
    }
}