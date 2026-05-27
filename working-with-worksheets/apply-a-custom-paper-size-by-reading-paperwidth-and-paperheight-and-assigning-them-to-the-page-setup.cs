using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Define custom paper dimensions (in inches)
        double paperWidth = 2.0;   // width in inches
        double paperHeight = 3.0;  // height in inches

        // Set the worksheet to use a custom paper size
        worksheet.PageSetup.PaperSize = PaperSizeType.Custom;
        worksheet.PageSetup.CustomPaperSize(paperWidth, paperHeight);

        // Save the workbook with the custom page setup
        workbook.Save("CustomPaperSizeOutput.xlsx");
    }
}