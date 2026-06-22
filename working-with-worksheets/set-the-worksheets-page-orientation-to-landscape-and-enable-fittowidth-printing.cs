using System;
using Aspose.Cells;

class SetPageOrientationAndFitToWidth
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set page orientation to Landscape
        worksheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Enable fit‑to‑width printing (fit to 1 page wide, height adjusts automatically)
        worksheet.PageSetup.FitToPagesWide = 1;
        worksheet.PageSetup.FitToPagesTall = 0; // 0 means auto‑adjust height

        // Save the workbook
        workbook.Save("PageSetupLandscapeFitToWidth.xlsx");
    }
}