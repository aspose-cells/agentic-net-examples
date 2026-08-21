// Title: Build a Pivot Table with Timeline and Export Column Chart to PDF using Aspose.Cells for .NET (C#)
// Description: C# code that creates a workbook, fills rows with Category, Date and Value, adds a pivot table (Category rows, Date page filter, Sum of Value), inserts a timeline linked to the Date field, generates a column chart from the source data, and saves the chart as a PDF (optionally keeping the Excel file).
// Keywords: Aspose.Cells | C# timeline chart | pivot table timeline | export chart to PDF | Aspose.Cells PDF export | column chart Aspose.Cells | timeline control .NET | automate Excel report | dynamic dashboard PDF | Aspose.Cells example
// Common Searches: Aspose.Cells add timeline to pivot table C# | Export Aspose.Cells chart as PDF | Create timeline control in Excel using Aspose.Cells | C# generate pivot table with date filter Aspose.Cells | How to save Aspose.Cells chart to PDF without workbook | Iterate rows and build timeline chart Aspose.Cells
// Developer Intent: Programmatically create a pivot‑based timeline and export its chart to a PDF file.
// Use Cases: Produce a sales‑by‑category column chart with an interactive timeline for periodic reporting. | Automate generation of Excel dashboards that include pivot tables and timelines, then deliver PDF snapshots to stakeholders. | Populate worksheet data from collections, attach a timeline filter, and create PDF exports for archival or email distribution.
// AI Prompts: Write C# code with Aspose.Cells that reads a list of objects, builds a pivot table with a Date page filter, adds a timeline control, and saves the resulting chart as a PDF. | Show how to iterate over rows, create a timeline‑enabled pivot table, customize the column chart, and export only the chart to PDF while keeping the workbook optional. | Explain steps to style the timeline, adjust chart layout, and generate a PDF report using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace AsposeCellsTimelineChartDemo
{
    // C# code that creates a workbook, fills rows with Category, Date and Value, adds a pivot table (Category rows, Date page filter, Sum of Value), inserts a timeline linked to the Date field, generates a column chart from the source data, and saves the chart as a PDF (optionally keeping the Excel file).
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate worksheet with sample data (Category, Date, Value)
                // Header row
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Date");
                cells["C1"].PutValue("Value");

                // Sample data rows
                string[] categories = { "Fruit", "Fruit", "Vegetable", "Vegetable", "Fruit" };
                DateTime[] dates = {
                    new DateTime(2023, 1, 1),
                    new DateTime(2023, 1, 2),
                    new DateTime(2023, 1, 1),
                    new DateTime(2023, 1, 2),
                    new DateTime(2023, 1, 3)
                };
                double[] values = { 120, 150, 80, 95, 200 };

                for (int i = 0; i < categories.Length; i++)
                {
                    int row = i + 2; // Data starts at row 2 (index 1)
                    cells[$"A{row}"].PutValue(categories[i]);
                    cells[$"B{row}"].PutValue(dates[i]);
                    cells[$"C{row}"].PutValue(values[i]);
                }

                // Define the data range for the pivot table (including header)
                string dataRange = "A1:C6";

                // Add a pivot table that will serve as the data source for the timeline
                int pivotIndex = sheet.PivotTables.Add(dataRange, "E3", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Configure pivot fields:
                // Row - Category, Page (filter) - Date, Data - Sum of Value
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Page, "Date"); // Required for timeline
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

                // Refresh pivot data to ensure it reflects the worksheet content
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a timeline control linked to the pivot table using the "Date" field
                // The timeline will be placed with its upper‑left corner at cell E1
                int timelineIndex = sheet.Timelines.Add(pivot, "E1", "Date");
                // Timeline timeline = sheet.Timelines[timelineIndex]; // Optional further customization

                // Create a column chart based on the original data (not the pivot)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 15, 0, 30, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("C2:C6", true);          // Values
                chart.NSeries.CategoryData = "A2:A6";     // Categories
                chart.Title.Text = "Sales by Category";

                // Export the chart to a PDF file
                string pdfPath = "TimelineChartOutput.pdf";
                chart.ToPdf(pdfPath);

                // Save the workbook (optional, to verify the timeline and chart in Excel)
                workbook.Save("TimelineChartDemo.xlsx");

                Console.WriteLine($"Chart exported to PDF: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
