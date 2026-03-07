using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartBookmarks
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -----------------------------------------------------------------
            // Prepare three worksheets, each containing a simple column chart
            // -----------------------------------------------------------------
            for (int i = 0; i < 3; i++)
            {
                // Get or create the worksheet
                Worksheet sheet = i == 0 ? workbook.Worksheets[0] : workbook.Worksheets.Add($"ChartSheet{i + 1}");

                // Add a title cell that will serve as the bookmark destination
                sheet.Cells["A1"].PutValue($"Chart Sheet {i + 1}");

                // Populate sample data for the chart
                sheet.Cells["A2"].PutValue("Category 1");
                sheet.Cells["A3"].PutValue("Category 2");
                sheet.Cells["A4"].PutValue("Category 3");
                sheet.Cells["B2"].PutValue(10 + i * 5);
                sheet.Cells["B3"].PutValue(20 + i * 5);
                sheet.Cells["B4"].PutValue(30 + i * 5);

                // Add a column chart covering the data range
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";
                chart.Title.Text = $"Sample Chart {i + 1}";
            }

            // ---------------------------------------------------------------
            // Build PDF bookmark hierarchy for the chart sheets
            // ---------------------------------------------------------------
            // Root bookmark (will appear as the top entry in the PDF outline)
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Workbook Charts",
                Destination = workbook.Worksheets[0].Cells["A1"], // Destination can be any cell; root uses first sheet
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            // Create a sub‑bookmark for each chart sheet
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                PdfBookmarkEntry sheetBookmark = new PdfBookmarkEntry
                {
                    Text = $"Chart Sheet {i + 1}",
                    Destination = workbook.Worksheets[i].Cells["A1"]
                };
                rootBookmark.SubEntry.Add(sheetBookmark);
            }

            // ---------------------------------------------------------------
            // Configure PDF save options with the bookmark structure
            // ---------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = rootBookmark,
                ExportDocumentStructure = true   // Preserve the bookmark hierarchy in the PDF
            };

            // Save the workbook as a PDF file with the defined bookmarks
            workbook.Save("ChartsWithBookmarks.pdf", pdfOptions);
        }
    }
}