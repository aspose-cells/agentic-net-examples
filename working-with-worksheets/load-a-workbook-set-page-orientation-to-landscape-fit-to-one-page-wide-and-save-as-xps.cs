using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set page orientation to Landscape
        sheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Fit the worksheet to one page wide (height adjusts automatically)
        sheet.PageSetup.FitToPagesWide = 1;
        sheet.PageSetup.FitToPagesTall = 0;

        // Create XPS save options (optional: force one page per sheet)
        XpsSaveOptions saveOptions = new XpsSaveOptions
        {
            OnePagePerSheet = true
        };

        // Save the workbook as XPS
        workbook.Save("output.xps", saveOptions);
    }
}