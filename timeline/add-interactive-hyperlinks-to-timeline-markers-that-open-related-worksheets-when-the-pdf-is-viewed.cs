using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsTimelineHyperlinkDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add two worksheets: one for the timeline, another as the hyperlink target
                Worksheet summarySheet = workbook.Worksheets[0];
                summarySheet.Name = "Summary";
                Worksheet detailSheet = workbook.Worksheets.Add("Details");

                // -------------------------------------------------
                // Populate data for a PivotTable on the Summary sheet
                // -------------------------------------------------
                Cells cells = summarySheet.Cells;
                cells["A1"].Value = "Date";
                cells["B1"].Value = "Category";
                cells["C1"].Value = "Amount";

                cells["A2"].Value = new DateTime(2023, 1, 1);
                cells["B2"].Value = "Food";
                cells["C2"].Value = 120;

                cells["A3"].Value = new DateTime(2023, 1, 5);
                cells["B3"].Value = "Travel";
                cells["C3"].Value = 300;

                cells["A4"].Value = new DateTime(2023, 2, 10);
                cells["B4"].Value = "Food";
                cells["C4"].Value = 80;

                // Create a PivotTable based on the data range
                int pivotIndex = summarySheet.PivotTables.Add("A1:C4", "E2", "SalesPivot");
                PivotTable pivot = summarySheet.PivotTables[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Column, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.RefreshData();
                pivot.CalculateData();

                // -------------------------------------------------
                // Add a Timeline linked to the PivotTable
                // -------------------------------------------------
                // Use the field index (0) for the Date field to avoid the "Error PivotField data" exception
                int timelineIndex = summarySheet.Timelines.Add(pivot, "G2", 0);
                Timeline timeline = summarySheet.Timelines[timelineIndex];

                // -------------------------------------------------
                // Add content to the Details sheet (the hyperlink target)
                // -------------------------------------------------
                detailSheet.Cells["A1"].Value = "Detail Information";
                detailSheet.Cells["A2"].Value = "This sheet is opened when the timeline marker is clicked.";

                // -------------------------------------------------
                // Create a hyperlink on the Timeline shape that points to Details!A1
                // -------------------------------------------------
                TimelineShape timelineShape = timeline.Shape;
                timelineShape.Hyperlink.Address = "'Details'!A1";
                timelineShape.Hyperlink.TextToDisplay = "Go to Details";

                // -------------------------------------------------
                // Save the workbook as PDF; the hyperlink on the timeline will be active in the PDF
                // -------------------------------------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions();
                workbook.Save("TimelineWithHyperlink.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}