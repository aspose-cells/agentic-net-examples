using System;
using Aspose.Cells;

class RemovePrintTitles
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through each worksheet and clear print title settings
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PageSetup pageSetup = sheet.PageSetup;

            // Remove rows and columns set to repeat on each printed page
            pageSetup.PrintTitleRows = "";
            pageSetup.PrintTitleColumns = "";
        }

        // Save the workbook with the changes applied
        workbook.Save("output.xlsx");
    }
}