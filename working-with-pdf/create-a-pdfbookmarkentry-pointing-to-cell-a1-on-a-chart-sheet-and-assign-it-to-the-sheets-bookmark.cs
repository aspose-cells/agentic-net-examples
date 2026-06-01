using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;   // Needed for PlacementType enum
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarkDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a worksheet that will host the chart
                Worksheet sheet = workbook.Worksheets.Add("MyChartSheet");

                // Populate data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.SetChartDataRange("A1:B4", true);

                // Set placement so the chart moves with cells
                chart.Placement = PlacementType.MoveAndSize;

                // Create a PDF bookmark entry that points to cell A1 on the worksheet
                PdfBookmarkEntry bookmarkEntry = new PdfBookmarkEntry
                {
                    Text = "ChartSheetStart",          // Title of the bookmark
                    Destination = sheet.Cells["A1"],   // Destination cell
                    IsOpen = true                      // Expand the bookmark by default
                };

                // Configure PDF save options with the bookmark
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Bookmark = bookmarkEntry
                };

                // Define output file path
                string outputPath = "ChartSheetBookmark.pdf";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a PDF with the bookmark
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}