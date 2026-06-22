using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data (optional)
        worksheet.Cells["A1"].PutValue("Sample");
        worksheet.Cells["B1"].PutValue("Data");

        // Configure page setup for unlimited width and height
        PageSetup pageSetup = worksheet.PageSetup;
        pageSetup.IsPercentScale = false;   // Ensure FitToPages properties are used
        pageSetup.FitToPagesWide = 0;       // 0 = unlimited pages wide
        pageSetup.FitToPagesTall = 0;       // 0 = unlimited pages tall

        // Save the workbook
        workbook.Save("UnlimitedFitPages.xlsx");
    }
}