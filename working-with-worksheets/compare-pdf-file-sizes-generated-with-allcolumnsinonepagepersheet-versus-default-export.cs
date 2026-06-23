using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfSizeComparison
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with many columns and rows to make the PDF sizable
            for (int row = 0; row < 50; row++)
            {
                for (int col = 0; col < 100; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Save PDF with default options
            string defaultPdfPath = "default.pdf";
            workbook.Save(defaultPdfPath);

            // Save PDF with AllColumnsInOnePagePerSheet enabled
            string allColumnsPdfPath = "allColumnsOnePage.pdf";
            PdfSaveOptions allColumnsOptions = new PdfSaveOptions();
            // Enable fitting all columns of a sheet onto a single page
            allColumnsOptions.AllColumnsInOnePagePerSheet = true;
            // Optional: also fit the whole sheet onto one page (height may be ignored)
            allColumnsOptions.OnePagePerSheet = true;
            workbook.Save(allColumnsPdfPath, allColumnsOptions);

            // Retrieve file sizes
            long defaultSize = new FileInfo(defaultPdfPath).Length;
            long allColumnsSize = new FileInfo(allColumnsPdfPath).Length;

            // Output the results
            Console.WriteLine($"Default PDF size: {defaultSize} bytes");
            Console.WriteLine($"AllColumnsInOnePagePerSheet PDF size: {allColumnsSize} bytes");
            Console.WriteLine($"Size difference: {Math.Abs(defaultSize - allColumnsSize)} bytes");
        }
    }
}