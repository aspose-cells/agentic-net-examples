using System;
using Aspose.Cells;

class SetMarginsExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set custom margins in centimeters
        worksheet.PageSetup.TopMargin = 2.0;      // Top margin: 2 cm
        worksheet.PageSetup.BottomMargin = 1.5;   // Bottom margin: 1.5 cm
        worksheet.PageSetup.LeftMargin = 1.0;     // Left margin: 1 cm
        worksheet.PageSetup.RightMargin = 1.0;    // Right margin: 1 cm

        // Save the workbook with the defined margins
        workbook.Save("CustomMargins.xlsx");
    }
}