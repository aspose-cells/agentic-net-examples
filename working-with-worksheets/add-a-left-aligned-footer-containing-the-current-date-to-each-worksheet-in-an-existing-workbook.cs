using System;
using Aspose.Cells;

class AddLeftFooterDate
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Set the left section (index 0) of the footer to the current date script
            sheet.PageSetup.SetFooter(0, "&D");
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}