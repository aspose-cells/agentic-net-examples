// Title: C# – Add a Data Series to an Existing Timeline Chart, Annotate, and Export to PDF with Aspose.Cells
// Description: This example shows how to load or create an Excel workbook, fill it with date‑sales data, build a pivot table, attach a timeline, create a column chart, add an extra data series, insert a textbox annotation, export the chart as a PDF, and save the updated workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# timeline | add chart series Aspose.Cells | export chart PDF Aspose.Cells | Excel timeline annotation | pivot table timeline .NET | Aspose.Cells chart textbox | C# generate PDF from chart | update existing timeline | Aspose.Cells workbook save | C# Excel automation
// Common Searches: Aspose.Cells add series to timeline chart C# | How to export Aspose.Cells chart to PDF | Create timeline from pivot table using Aspose.Cells | Add textbox annotation to Excel chart with Aspose.Cells | Update existing Excel timeline programmatically | C# generate PDF report from Excel chart
// Developer Intent: Add a new series to a timeline‑driven chart, place a textbox annotation, and produce a PDF version of the chart.
// Use Cases: Enhance a sales dashboard by adding a comparative series to a timeline chart. | Provide contextual notes on the worksheet via a textbox positioned near the chart. | Automate creation of PDF charts for inclusion in reports or email distribution. | Preserve the modified workbook for further analysis or future updates.
// AI Prompts: Generate C# code with Aspose.Cells that loads an Excel file, creates a pivot table, adds a timeline, inserts a column chart, adds an extra data series, places a textbox annotation, and exports the chart to PDF. | Show how to implement robust error handling when loading a workbook, adding a timeline, and exporting a chart using Aspose.Cells. | Explain the steps to refresh a pivot‑based timeline after adding new data in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsTimelineUpdate
{
    // This example shows how to load or create an Excel workbook, fill it with date‑sales data, build a pivot table, attach a timeline, create a column chart, add an extra data series, insert a textbox annotation, export the chart as a PDF, and save the updated workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            Workbook workbook;

            // Load existing workbook if present; otherwise create a new one.
            try
            {
                if (File.Exists("ExistingTimeline.xlsx"))
                {
                    workbook = new Workbook("ExistingTimeline.xlsx");
                }
                else
                {
                    workbook = new Workbook();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                workbook = new Workbook();
            }

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Ensure there is sample data for a pivot table (Date & Sales).
            // ------------------------------------------------------------
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Sales");
            cells["A2"].PutValue(new DateTime(2023, 1, 1));
            cells["B2"].PutValue(1200);
            cells["A3"].PutValue(new DateTime(2023, 2, 1));
            cells["B3"].PutValue(1500);
            cells["A4"].PutValue(new DateTime(2023, 3, 1));
            cells["B4"].PutValue(1800);
            cells["A5"].PutValue(new DateTime(2023, 4, 1));
            cells["B5"].PutValue(2100);
            cells["A6"].PutValue(new DateTime(2023, 5, 1));
            cells["B6"].PutValue(2400);
            cells["A7"].PutValue(new DateTime(2023, 6, 1));
            cells["B7"].PutValue(2700);

            // ------------------------------------------------------------
            // Create a pivot table that will serve as the data source for the timeline.
            // ------------------------------------------------------------
            int pivotIdx = sheet.PivotTables.Add("A1:B7", "D1", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.RefreshData();
            pivot.CalculateData();

            // ------------------------------------------------------------
            // Add (or update) a timeline linked to the pivot table.
            // ------------------------------------------------------------
            try
            {
                // Use zero‑based row/column indices (F1 => row 0, column 5)
                sheet.Timelines.Add(pivot, 0, 5, "Date");
                // Note: Setting Caption/Name is omitted due to API version constraints.
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error adding timeline: {ex.Message}");
            }

            // ------------------------------------------------------------
            // Create a chart that visualizes the same data.
            // ------------------------------------------------------------
            int chartIdx = sheet.Charts.Add(ChartType.Column, 10, 0, 25, 15);
            Chart chart = sheet.Charts[chartIdx];

            // First series (existing data range).
            chart.NSeries.Add("B2:B7", true);
            chart.NSeries.CategoryData = "A2:A7";

            // ------------------------------------------------------------
            // Add a new data series to the chart.
            // ------------------------------------------------------------
            chart.NSeries.Add("B3:B7", true);

            // ------------------------------------------------------------
            // Add an annotation (textbox) to the worksheet near the chart.
            // ------------------------------------------------------------
            Shape annotation = sheet.Shapes.AddTextBox(5, 0, 5, 0, 200, 50);
            annotation.Text = "Updated with new series";
            annotation.Font.Size = 12;
            annotation.Font.Color = Color.Blue;

            // ------------------------------------------------------------
            // Export the chart to a PDF file.
            // ------------------------------------------------------------
            try
            {
                chart.ToPdf("SalesTimelineChart.pdf");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error exporting chart to PDF: {ex.Message}");
            }

            // ------------------------------------------------------------
            // Save the workbook with the updated timeline and chart.
            // ------------------------------------------------------------
            try
            {
                workbook.Save("UpdatedTimelineWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}
