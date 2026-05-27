using System;
using Aspose.Cells;

class SetHeadersBatch
{
    static void Main()
    {
        // Load the existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Define identical header scripts for all worksheets
        string leftHeader   = "&\"Arial,Bold\"&12My Company";
        string centerHeader = "&\"Arial\"&12Report Title";
        string rightHeader  = "&D &T"; // Date and Time

        // Apply the headers to every worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.SetHeader(0, leftHeader);   // Left section
            pageSetup.SetHeader(1, centerHeader); // Center section
            pageSetup.SetHeader(2, rightHeader);  // Right section
        }

        // Save the modified workbook (replace with desired output path)
        workbook.Save("output.xlsx");
    }
}