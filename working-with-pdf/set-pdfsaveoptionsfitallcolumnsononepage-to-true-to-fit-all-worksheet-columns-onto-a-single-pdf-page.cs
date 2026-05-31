using System;
using Aspose.Cells;

namespace AsposeCellsFitAllColumnsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data spanning many columns
            for (int col = 0; col < 30; col++)
            {
                // Header
                sheet.Cells[0, col].PutValue($"Header {col + 1}");
                // Sample data rows
                for (int row = 1; row <= 5; row++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col + 1}");
                }
            }

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Fit all columns of each sheet onto a single PDF page
            pdfOptions.AllColumnsInOnePagePerSheet = true;

            // Save the workbook as PDF with the specified options
            workbook.Save("AllColumnsOnePage.pdf", pdfOptions);
        }
    }
}