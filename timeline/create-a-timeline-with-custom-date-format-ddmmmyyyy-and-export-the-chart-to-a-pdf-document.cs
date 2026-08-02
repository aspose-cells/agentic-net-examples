// Title: C# Aspose.Cells: Timeline with dd‑MMM‑yyyy Format and Export Chart to PDF
// Description: Creates a workbook, adds dates and values, applies the custom format dd‑MMM‑yyyy, builds a pivot table, inserts a timeline, draws a line chart, and saves the chart as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells timeline | custom date format dd-MMM-yyyy | C# chart to PDF | Aspose.Cells pivot table | export chart as PDF | .NET timeline chart | line chart X‑axis date format | Aspose.Cells PDF export | timeline filter Aspose | C# Aspose.Cells example
// Common Searches: Aspose.Cells timeline custom date format C# | Export Aspose.Cells chart to PDF | Create timeline linked to pivot table Aspose.Cells | Set X axis date format Aspose.Cells chart | C# example timeline PDF export
// Developer Intent: Generate a timeline tied to a pivot table, format its dates as dd‑MMM‑yyyy, and export the associated chart to a PDF file.
// Use Cases: Financial dashboards where the timeline shows dates like 01‑Jan‑2021 and the chart is shared as a PDF report. | Project schedule visualizations with a timeline filter, custom date labels, and printable PDF output for stakeholders.
// AI Prompts: Write C# code using Aspose.Cells to create a timeline with the date format dd‑MMM‑yyyy, link it to a pivot table, and export the chart to PDF. | Explain how to set the X‑axis date format for a line chart in Aspose.Cells and use TimelineCollection for data filtering. | Give best practices for handling exceptions when converting an Aspose.Cells chart to a PDF document.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

// Creates a workbook, adds dates and values, applies the custom format dd‑MMM‑yyyy, builds a pivot table, inserts a timeline, draws a line chart, and saves the chart as a PDF using Aspose.Cells for .NET.
public class TimelineChartPdfDemo
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unhandled exception: " + ex.Message);
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate worksheet with dates and values
        cells["A1"].PutValue("Date");
        cells["B1"].PutValue("Value");

        DateTime[] dates = new DateTime[]
        {
            new DateTime(2021, 1, 1),
            new DateTime(2021, 2, 1),
            new DateTime(2021, 3, 1),
            new DateTime(2021, 4, 1)
        };
        double[] values = { 10, 20, 30, 40 };

        for (int i = 0; i < dates.Length; i++)
        {
            cells[i + 1, 0].PutValue(dates[i]);   // Column A
            cells[i + 1, 1].PutValue(values[i]); // Column B
        }

        // Apply custom date format "dd-MMM-yyyy" to the date column
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "dd-MMM-yyyy";
        for (int i = 1; i <= dates.Length; i++)
        {
            cells[i, 0].SetStyle(dateStyle);
        }

        // Create a pivot table (required as a data source for the timeline)
        PivotTableCollection pivots = sheet.PivotTables;
        int pivotIndex = pivots.Add("A1:B5", "D1", "Pivot1");
        PivotTable pivot = pivots[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a timeline linked to the pivot table using the "Date" field
        TimelineCollection timelines = sheet.Timelines;
        timelines.Add(pivot, "F1", "Date");

        // Create a line chart that uses the same data range
        int chartIndex = sheet.Charts.Add(ChartType.Line, 10, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B5", true);          // Values
        chart.NSeries.CategoryData = "A2:A5";      // Dates

        // Set the X‑axis (category) values format to the custom date format
        chart.NSeries[0].XValuesFormatCode = "dd-MMM-yyyy";

        // Export the chart to a PDF file
        string outputPath = "TimelineChart.pdf";
        try
        {
            chart.ToPdf(outputPath);
            Console.WriteLine($"Chart exported to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to export PDF: " + ex.Message);
        }
    }
}
