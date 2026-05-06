using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarkDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and populate data for three charts
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();

            // ----- Worksheet 1: Sales Data -----
            Worksheet salesSheet = workbook.Worksheets[0];
            salesSheet.Name = "Sales";
            salesSheet.Cells["A1"].PutValue("Month");
            salesSheet.Cells["B1"].PutValue("Sales");
            salesSheet.Cells["A2"].PutValue("Jan");
            salesSheet.Cells["B2"].PutValue(120);
            salesSheet.Cells["A3"].PutValue("Feb");
            salesSheet.Cells["B3"].PutValue(150);
            salesSheet.Cells["A4"].PutValue("Mar");
            salesSheet.Cells["B4"].PutValue(180);

            // Add a column chart for sales
            int salesChartIdx = salesSheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart salesChart = salesSheet.Charts[salesChartIdx];
            salesChart.NSeries.Add("B2:B4", true);
            salesChart.NSeries.CategoryData = "A2:A4";
            salesChart.Title.Text = "Quarterly Sales";

            // ----- Worksheet 2: Profit Data -----
            Worksheet profitSheet = workbook.Worksheets.Add("Profit");
            profitSheet.Cells["A1"].PutValue("Month");
            profitSheet.Cells["B1"].PutValue("Profit");
            profitSheet.Cells["A2"].PutValue("Jan");
            profitSheet.Cells["B2"].PutValue(30);
            profitSheet.Cells["A3"].PutValue("Feb");
            profitSheet.Cells["B3"].PutValue(45);
            profitSheet.Cells["A4"].PutValue("Mar");
            profitSheet.Cells["B4"].PutValue(55);

            // Add a line chart for profit
            int profitChartIdx = profitSheet.Charts.Add(ChartType.Line, 5, 0, 20, 12);
            Chart profitChart = profitSheet.Charts[profitChartIdx];
            profitChart.NSeries.Add("B2:B4", true);
            profitChart.NSeries.CategoryData = "A2:A4";
            profitChart.Title.Text = "Quarterly Profit";

            // ----- Worksheet 3: Growth Data -----
            Worksheet growthSheet = workbook.Worksheets.Add("Growth");
            growthSheet.Cells["A1"].PutValue("Month");
            growthSheet.Cells["B1"].PutValue("Growth %");
            growthSheet.Cells["A2"].PutValue("Jan");
            growthSheet.Cells["B2"].PutValue(5);
            growthSheet.Cells["A3"].PutValue("Feb");
            growthSheet.Cells["B3"].PutValue(7);
            growthSheet.Cells["A4"].PutValue("Mar");
            growthSheet.Cells["B4"].PutValue(9);

            // Add a pie chart for growth distribution
            int growthChartIdx = growthSheet.Charts.Add(ChartType.Pie, 5, 0, 20, 12);
            Chart growthChart = growthSheet.Charts[growthChartIdx];
            growthChart.NSeries.Add("B2:B4", true);
            growthChart.NSeries.CategoryData = "A2:A4";
            growthChart.Title.Text = "Growth Distribution";

            // ------------------------------------------------------------
            // 2. Create PDF bookmark hierarchy for the chart sheets
            // ------------------------------------------------------------
            // Root bookmark representing the whole report
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Quarterly Report",
                Destination = salesSheet.Cells["A1"], // Destination can be any cell; using first sheet as entry point
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            // Sub‑bookmark for Sales chart
            PdfBookmarkEntry salesBookmark = new PdfBookmarkEntry
            {
                Text = "Sales Chart",
                Destination = salesSheet.Cells["A1"]
            };

            // Sub‑bookmark for Profit chart
            PdfBookmarkEntry profitBookmark = new PdfBookmarkEntry
            {
                Text = "Profit Chart",
                Destination = profitSheet.Cells["A1"]
            };

            // Sub‑bookmark for Growth chart
            PdfBookmarkEntry growthBookmark = new PdfBookmarkEntry
            {
                Text = "Growth Chart",
                Destination = growthSheet.Cells["A1"]
            };

            // Assemble hierarchy
            rootBookmark.SubEntry.Add(salesBookmark);
            rootBookmark.SubEntry.Add(profitBookmark);
            rootBookmark.SubEntry.Add(growthBookmark);

            // ------------------------------------------------------------
            // 3. Configure PDF save options
            // ------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Attach the bookmark hierarchy to the PDF
                Bookmark = rootBookmark,

                // Export document structure so that bookmarks are recognized by PDF readers
                ExportDocumentStructure = true,

                // Optional: one page per sheet to keep each chart on its own page
                OnePagePerSheet = true
            };

            // ------------------------------------------------------------
            // 4. Save the workbook as a PDF with bookmarks
            // ------------------------------------------------------------
            // The resulting PDF will contain three pages (one per chart) and a
            // bookmark pane that allows users to jump directly to each chart.
            // This is useful in scenarios such as:
            //   • Interactive financial reports where executives can navigate
            //     quickly to the chart of interest.
            //   • Automated report generation where a table of contents is
            //     represented by PDF bookmarks.
            //   • Embedding charts in e‑books or documentation with easy
            //     navigation.
            workbook.Save("QuarterlyReport.pdf", pdfOptions);
        }
    }
}