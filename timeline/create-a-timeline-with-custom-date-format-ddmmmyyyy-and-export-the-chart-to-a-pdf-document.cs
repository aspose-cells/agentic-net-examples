// Title: Aspose.Cells for .NET: Create a Timeline with dd‑MMM‑yyyy Format and Export Its Chart to PDF (C#)
// Description: A C# sample that builds a workbook, applies the custom date pattern dd‑MMM‑yyyy to cells, creates a PivotTable, adds a Timeline bound to the Date field, generates a line chart using the same data, formats the chart’s X‑axis with the custom date pattern, and exports the chart directly to a PDF file.
// Keywords: Aspose.Cells | C# | Timeline | custom date format | dd-MMM-yyyy | PivotTable | chart to PDF | export chart PDF | Aspose.Cells timeline | Aspose.Cells chart PDF | Aspose.Cells example
// Common Searches: Aspose.Cells timeline custom date format | Export chart to PDF Aspose.Cells C# | Create timeline from pivot table Aspose.Cells | Set X axis date format Aspose chart | C# Aspose.Cells generate PDF from chart
// Developer Intent: Generate a timeline with dd‑MMM‑yyyy dates, link it to a pivot table, create a line chart, and save the chart as a PDF using Aspose.Cells for .NET.
// Use Cases: Interactive sales dashboard where a timeline filters data and the chart is exported as a PDF for periodic reporting. | Automated financial statements that require uniform date formatting and PDF chart snapshots for distribution. | Producing printable chart PDFs from Excel workbooks for client presentations or archival purposes.
// AI Prompts: Write C# code with Aspose.Cells to create a timeline using the date format dd‑MMM‑yyyy, bind it to a pivot table, add a line chart, and export the chart to a PDF file. | Explain how to set a custom X‑axis date format for a chart in Aspose.Cells and connect the chart to a timeline for synchronized filtering. | Provide step‑by‑step instructions to generate a PDF of a chart that reflects the selected date range of a timeline using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelinePdfDemo
{
    // A C# sample that builds a workbook, applies the custom date pattern dd‑MMM‑yyyy to cells, creates a PivotTable, adds a Timeline bound to the Date field, generates a line chart using the same data, formats the chart’s X‑axis with the custom date pattern, and exports the chart directly to a PDF file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ----- Populate worksheet with sample date and value data -----
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Value");

            DateTime[] dates = {
                new DateTime(2021, 1, 5),
                new DateTime(2021, 2, 12),
                new DateTime(2021, 3, 20),
                new DateTime(2021, 4, 15)
            };
            int[] values = { 100, 150, 130, 170 };

            for (int i = 0; i < dates.Length; i++)
            {
                // Row index starts at 1 because row 0 holds headers
                int row = i + 1;
                cells[$"A{row + 1}"].PutValue(dates[i]);   // A2, A3, ...
                cells[$"B{row + 1}"].PutValue(values[i]); // B2, B3, ...

                // Apply custom date format "dd-MMM-yyyy" to the date cells
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Custom = "dd-MMM-yyyy";
                cells[$"A{row + 1}"].SetStyle(dateStyle);
            }

            // ----- Create a PivotTable (required as data source for Timeline) -----
            // Define the source range (including headers)
            string sourceRange = "A1:B5";
            // Destination cell for the pivot table
            string pivotDest = "D1";
            int pivotIndex = sheet.PivotTables.Add(sourceRange, pivotDest, "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh and calculate the pivot data
            pivot.RefreshData();
            pivot.CalculateData();

            // ----- Add a Timeline linked to the PivotTable -----
            // Place the Timeline starting at cell E1 and bind it to the "Date" field
            int timelineIndex = sheet.Timelines.Add(pivot, "E1", "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];
            timeline.Caption = "Sales Timeline";

            // ----- Create a chart that uses the same data -----
            int chartIndex = sheet.Charts.Add(ChartType.Line, 10, 0, 25, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B5", true);          // Values
            chart.NSeries.CategoryData = "A2:A5";     // Dates

            // Ensure the X‑axis (category axis) displays dates with the custom format
            // This can be done by setting the XValues format code directly
            chart.NSeries[0].XValuesFormatCode = "dd-MMM-yyyy";

            // Optional: give the chart a title
            chart.Title.Text = "Monthly Sales";

            // ----- Export the chart to a PDF file -----
            // The ToPdf method is part of Aspose.Cells.Charts.Chart
            chart.ToPdf("TimelineChart.pdf");

            // (Optional) Save the workbook to verify the timeline and chart in Excel format
            workbook.Save("TimelineWithChart.xlsx");
        }
    }
}
