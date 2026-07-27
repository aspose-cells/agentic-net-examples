// Title: C# – Build a Pivot‑Table Timeline, Create a Column Chart, and Export to PDF with Aspose.Cells
// Description: Shows how to programmatically create a workbook, populate it with dates and values, add a pivot table, attach a timeline control to the Date field, generate a column chart from the pivot data, and save the chart as a PDF (with optional workbook save) using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# timeline | pivot table timeline Aspose.Cells | export chart to PDF .NET | column chart from pivot Aspose.Cells | Aspose.Cells PDF export example | timeline control worksheet | Aspose.Cells API chart PDF | C# generate timeline entries | Aspose.Cells workbook save
// Common Searches: how to add a timeline to a pivot table using Aspose.Cells | export a chart created from a pivot table to PDF in C# | populate worksheet rows with dates and values for a timeline chart | Aspose.Cells example for timeline control and PDF export | C# code to create a column chart from pivot data and save as PDF
// Developer Intent: Create a timeline linked to a pivot table, build a column chart from that data, and export the chart as a PDF file using Aspose.Cells for .NET.
// Use Cases: Enable interactive date filtering on a pivot‑driven chart via a timeline control. | Generate PDF reports that contain a column chart reflecting time‑series data. | Automate workbook creation with timelines and charts for distribution to stakeholders.
// AI Prompts: Provide C# code to add multiple timelines for different date fields with Aspose.Cells. | Show how to export the entire worksheet, including timelines, to a single PDF file. | Explain how to bind a chart series directly to a pivot table instead of a raw range in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineChartDemo
{
    // Shows how to programmatically create a workbook, populate it with dates and values, add a pivot table, attach a timeline control to the Date field, generate a column chart from the pivot data, and save the chart as a PDF (with optional workbook save) using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ------------------------------------------------------------
                // 2. Populate sample data:
                //    Column A -> "Date"
                //    Column B -> "Value"
                //    Rows 2..6 contain sample dates and numeric values
                // ------------------------------------------------------------
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Value");

                DateTime startDate = new DateTime(2023, 1, 1);
                for (int i = 0; i < 5; i++)
                {
                    cells[i + 1, 0].PutValue(startDate.AddDays(i));   // Date column
                    cells[i + 1, 1].PutValue(100 + i * 50);           // Value column
                }

                // ------------------------------------------------------------
                // 3. Create a PivotTable that will serve as the data source for timelines
                //    The pivot will use the whole data range (A1:B6) and be placed at C1
                // ------------------------------------------------------------
                int pivotIndex = sheet.PivotTables.Add("A1:B6", "C1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Place the Date field in the Page area (required for timelines)
                pivot.AddFieldToArea(PivotFieldType.Page, "Date");
                // Place the Value field in the Data area
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");

                // Refresh and calculate the pivot data
                pivot.RefreshData();
                pivot.CalculateData();

                // ------------------------------------------------------------
                // 4. Add timelines linked to the pivot table
                // ------------------------------------------------------------
                int timelineStartRow = 0;   // top-left corner row index for the first timeline
                int timelineStartCol = 5;   // start column index (F column)

                // Add a single timeline (multiple timelines are unnecessary for the same field)
                sheet.Timelines.Add(pivot, timelineStartRow, timelineStartCol, "Date");

                // ------------------------------------------------------------
                // 5. Create a chart that visualizes the pivot data.
                //    The chart will be placed below the timelines.
                // ------------------------------------------------------------
                int chartRow1 = 10;
                int chartCol1 = 0;
                int chartRow2 = 25;
                int chartCol2 = 8;

                int chartIndex = sheet.Charts.Add(ChartType.Column, chartRow1, chartCol1, chartRow2, chartCol2);
                Chart chart = sheet.Charts[chartIndex];

                // Use the original data range for the chart series
                chart.NSeries.Add("B2:B6", true);
                chart.NSeries.CategoryData = "A2:A6";
                chart.Title.Text = "Sample Values Over Time";

                // ------------------------------------------------------------
                // 6. Export the chart to a PDF file.
                // ------------------------------------------------------------
                string pdfPath = "TimelineChartOutput.pdf";
                chart.ToPdf(pdfPath);

                // ------------------------------------------------------------
                // 7. Save the workbook (optional, to verify the timelines and chart)
                // ------------------------------------------------------------
                string workbookPath = "TimelineChartDemo.xlsx";
                workbook.Save(workbookPath);

                Console.WriteLine("Chart exported to PDF and workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
