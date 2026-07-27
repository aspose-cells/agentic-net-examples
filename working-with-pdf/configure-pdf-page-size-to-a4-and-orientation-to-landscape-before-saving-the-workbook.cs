using System;
using Aspose.Cells;

namespace AsposeCellsPdfDemo
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the paper size for the whole workbook to A4
            workbook.Settings.PaperSize = PaperSizeType.PaperA4;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set page orientation to Landscape
            sheet.PageSetup.Orientation = PageOrientationType.Landscape;

            // (Optional) Add some content to demonstrate the PDF output
            sheet.Cells["A1"].PutValue("Sample PDF with A4 Landscape");

            // Save the workbook as PDF
            workbook.Save("Output_A4_Landscape.pdf", SaveFormat.Pdf);
        }
    }
}