// Title: Add Data Series to Pivot Timeline Chart, Annotate, and Export PDF – Aspose.Cells C#
// Description: Creates a workbook with date‑based sales data, builds a pivot table, adds a timeline linked to the Date field, generates a column chart, appends a second sales series, inserts a textbox note, and exports the updated chart to PDF while saving the workbook.
// Keywords: Aspose.Cells C# add chart series | pivot timeline Aspose.Cells | export chart to PDF .NET | worksheet textbox annotation Aspose.Cells | update chart data range programmatically | timeline linked to pivot table | column chart from pivot data
// Common Searches: how to add another series to a pivot chart using Aspose.Cells | export timeline chart as PDF with Aspose.Cells .NET | add textbox annotation to Excel worksheet programmatically | create timeline for pivot table when Timeline class missing | Aspose.Cells chart to PDF example
// Developer Intent: Add a new data series to an existing timeline‑linked chart, place a descriptive annotation, and generate a PDF of the updated chart using Aspose.Cells for .NET.
// Use Cases: Generate a PDF report that shows multiple sales series on a timeline‑driven column chart. | Provide a visual note on the worksheet to indicate chart updates or data changes. | Support environments where the Aspose.Cells Timeline class is unavailable by using the generic timeline API.
// AI Prompts: Write C# code with Aspose.Cells to add a second series to a pivot‑based column chart, set its series name, and export the chart to PDF with a textbox annotation. | Explain how to create a timeline linked to a pivot table's date field and handle cases where the Timeline class is not present in the library. | Show how to programmatically adjust a chart's data source ranges when new columns are added to the worksheet.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook with date‑based sales data, builds a pivot table, adds a timeline linked to the Date field, generates a column chart, appends a second sales series, inserts a textbox note, and exports the updated chart to PDF while saving the workbook.
class UpdateTimelineAndExportPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data with dates and two sales series
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Sales1");
            cells["C1"].PutValue("Sales2");
            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 5; i++)
            {
                cells[i + 1, 0].PutValue(startDate.AddMonths(i));
                cells[i + 1, 1].PutValue(100 + i * 10);   // Sales1 values
                cells[i + 1, 2].PutValue(150 + i * 15);   // Sales2 values
            }

            // Create a pivot table using Date as row field and Sales1 as data field
            int pivotIdx = sheet.PivotTables.Add("A1:C6", "E1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales1");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the Date field of the pivot table (if supported)
            // The Timeline class may not be available in older Aspose.Cells versions.
            // Therefore, we add the timeline without using the Timeline type.
            int timelineIdx = sheet.Timelines.Add(pivot, "G1", "Date");
            // If the Timeline object is available, you could set its caption like this:
            // sheet.Timelines[timelineIdx].Caption = "Sales Timeline";

            // Add a column chart based on the pivot data (Sales1)
            int chartIdx = sheet.Charts.Add(ChartType.Column, 10, 0, 25, 15);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("=Sheet1!B2:B6", true);               // Series for Sales1
            chart.NSeries.CategoryData = "=Sheet1!A2:A6";          // Categories (dates)

            // Add a new data series (Sales2) to the existing chart
            chart.NSeries.Add("=Sheet1!C2:C6", true);
            chart.NSeries[1].Name = "Sales2";

            // Add a textbox annotation to the worksheet (acts as a visual note)
            Shape annotation = sheet.Shapes.AddTextBox(5, 0, 5, 0, 200, 30);
            annotation.Text = "Updated with new series";
            annotation.Font.Size = 12;
            annotation.Font.Color = Color.Blue;

            // Export the chart (which reflects the updated data) to PDF
            chart.ToPdf("UpdatedTimelineChart.pdf");

            // Save the workbook for reference
            workbook.Save("UpdatedTimelineWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
