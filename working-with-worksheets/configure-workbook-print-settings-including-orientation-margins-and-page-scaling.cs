using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the PageSetup object for the worksheet
        PageSetup pageSetup = worksheet.PageSetup;

        // Set page orientation to Landscape
        pageSetup.Orientation = PageOrientationType.Landscape;

        // Set margins (values are in centimeters)
        pageSetup.TopMargin = 1.5;      // Top margin
        pageSetup.BottomMargin = 1.5;   // Bottom margin
        pageSetup.LeftMargin = 2.0;     // Left margin
        pageSetup.RightMargin = 2.0;    // Right margin

        // Set page scaling: fit the worksheet to 1 page wide and 1 page tall
        pageSetup.FitToPagesWide = 1;
        pageSetup.FitToPagesTall = 1;

        // Alternatively, you could use a percentage zoom
        // pageSetup.Zoom = 80; // 80%

        // Save the workbook as an XLSX file (lifecycle rule: save)
        workbook.Save("PrintSettingsDemo.xlsx");
    }
}