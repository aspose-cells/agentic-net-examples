using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

namespace AsposeCellsTimelineHyperlinkDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Data";

                // Populate sample data for a pivot table
                sheet1.Cells["A1"].Value = "Date";
                sheet1.Cells["B1"].Value = "Category";
                sheet1.Cells["C1"].Value = "Sales";

                sheet1.Cells["A2"].Value = new DateTime(2023, 1, 1);
                sheet1.Cells["B2"].Value = "A";
                sheet1.Cells["C2"].Value = 100;

                sheet1.Cells["A3"].Value = new DateTime(2023, 2, 1);
                sheet1.Cells["B3"].Value = "B";
                sheet1.Cells["C3"].Value = 150;

                sheet1.Cells["A4"].Value = new DateTime(2023, 3, 1);
                sheet1.Cells["B4"].Value = "A";
                sheet1.Cells["C4"].Value = 200;

                // Create a pivot table based on the data
                int pivotIdx = sheet1.PivotTables.Add("A1:C4", "E1", "SalesPivot");
                PivotTable pivot = sheet1.PivotTables[pivotIdx];

                // Add the Date field to the Page (filter) area – required for a timeline
                pivot.AddFieldToArea(PivotFieldType.Page, "Date");
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh and calculate the pivot table
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a timeline control linked to the Date field of the pivot table
                int timelineIdx = sheet1.Timelines.Add(pivot, "A10", "Date");
                Timeline timeline = sheet1.Timelines[timelineIdx];

                // Add a second worksheet that will be the hyperlink target
                Worksheet sheet2 = workbook.Worksheets.Add("Details");
                sheet2.Cells["A1"].Value = "Details for selected period";

                // Set a hyperlink on the timeline shape to open the target worksheet cell when PDF is viewed
                timeline.Shape.Hyperlink.Address = "Details!A1";
                timeline.Shape.Hyperlink.TextToDisplay = "Go to Details";

                // Save the workbook as PDF with the hyperlink active
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