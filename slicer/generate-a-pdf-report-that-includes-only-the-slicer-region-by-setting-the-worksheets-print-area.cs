using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace SlicerPdfReport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data (replace with actual slicer data as needed)
            sheet.Cells["A1"].PutValue("Header");
            for (int i = 2; i <= 20; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Define the range that contains the slicer (example: C5:D10)
            // Set this range as the print area so only this region is exported
            sheet.PageSetup.PrintArea = "C5:D10";

            // Save the workbook as PDF; the print area limits the exported content
            workbook.Save("SlicerRegionReport.pdf", SaveFormat.Pdf);
        }
    }
}