// Title: Create a PivotTable timeline and export the workbook to a single PDF with the timeline repeated on every page using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add timeline to a pivot table and keep it on each PDF page | How to print a timeline control on every page when saving a workbook as PDF with Aspose.Cells | C# export workbook to PDF with a repeating timeline using PdfSaveOptions | Set print title rows for a timeline in Aspose.Cells before PDF conversion | Aspose.Cells .NET generate PDF with pivot table timeline across multiple pages
// Tags: pivot table timeline Aspose.Cells | export workbook to PDF Aspose.Cells | repeat timeline on printed pages | PdfSaveOptions OnePagePerSheet false | print title rows timeline Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

namespace TimelinePdfDemo
{
    // The example creates a workbook, populates it with sample sales data, builds a PivotTable, adds a Timeline linked to the PivotTable, configures page setup to repeat the first rows (including the timeline) on every printed page, and saves the entire workbook as a single PDF using PdfSaveOptions with OnePagePerSheet disabled.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output PDF file
            const string outputPdfPath = "WorkbookWithTimeline.pdf";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // 1. Populate sample data (including a date field for the timeline)
            // ------------------------------------------------------------
            sheet.Cells["A1"].PutValue("Ship Date");
            sheet.Cells["B1"].PutValue("Region");
            sheet.Cells["C1"].PutValue("Sales");

            DateTime[] dates = {
                new DateTime(2023, 1, 5),
                new DateTime(2023, 1, 12),
                new DateTime(2023, 2, 3),
                new DateTime(2023, 2, 20),
                new DateTime(2023, 3, 8),
                new DateTime(2023, 3, 15)
            };

            string[] regions = { "North", "South", "East", "West", "North", "South" };
            double[] sales = { 1200, 850, 960, 1430, 1100, 970 };

            for (int i = 0; i < dates.Length; i++)
            {
                int row = i + 2; // data starts at row 2
                sheet.Cells[$"A{row}"].PutValue(dates[i]);
                sheet.Cells[$"B{row}"].PutValue(regions[i]);
                sheet.Cells[$"C{row}"].PutValue(sales[i]);
            }

            // ------------------------------------------------------------
            // 2. Create a PivotTable based on the data
            // ------------------------------------------------------------
            string sourceRange = "A1:C7";
            string pivotDestination = "E3";

            int pivotIndex = sheet.PivotTables.Add(sourceRange, pivotDestination, "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields – Ship Date must be a row field for the timeline to work
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Ship Date");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Apply a visual style
            pivotTable.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

            // ------------------------------------------------------------
            // 3. Add a Timeline control linked to the PivotTable
            // ------------------------------------------------------------
            try
            {
                // Add the timeline; row and column specify the top‑left cell where the timeline will be placed
                int timelineIndex = sheet.Timelines.Add(pivotTable, 0, 6, "Ship Date");
                Timeline timeline = sheet.Timelines[timelineIndex];

                // Optional size adjustments (not required for PDF export)
                // timeline.Width = 400;
                // timeline.Height = 30;
                // timeline.IsPrintObject = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error adding timeline: {ex.Message}");
                // Continue without a timeline if it cannot be added
            }

            // ------------------------------------------------------------
            // 4. Configure page setup so the timeline appears on every printed page
            // ------------------------------------------------------------
            sheet.PageSetup.PrintTitleRows = "$1:$2";
            sheet.PageSetup.PrintGridlines = true;
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 0; // automatic height

            // ------------------------------------------------------------
            // 5. Save the entire workbook to a single PDF file
            // ------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = false,
                Compliance = PdfCompliance.PdfA1b
            };

            try
            {
                workbook.Save(outputPdfPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved to PDF: {outputPdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while saving PDF: {ex.Message}");
            }
        }
    }
}
