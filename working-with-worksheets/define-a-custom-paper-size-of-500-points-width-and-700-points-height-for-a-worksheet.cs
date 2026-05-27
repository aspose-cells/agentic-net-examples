using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Convert points to inches (1 point = 1/72 inch)
        double widthInches = 500.0 / 72.0;   // 500 points width
        double heightInches = 700.0 / 72.0;  // 700 points height

        // Set the custom paper size for the worksheet
        worksheet.PageSetup.CustomPaperSize(widthInches, heightInches);
        // Explicitly set the paper size type to Custom
        worksheet.PageSetup.PaperSize = PaperSizeType.Custom;

        // Save the workbook
        workbook.Save("CustomPaperSize.xlsx");
    }
}