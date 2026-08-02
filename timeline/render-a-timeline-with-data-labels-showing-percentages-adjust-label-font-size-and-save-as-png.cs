// Title: Render a PivotTable Timeline with a Pie Chart (percentage labels, custom font) and export as PNG using Aspose.Cells for .NET
// Description: Creates a workbook, adds category, value and date data, builds a PivotTable with a date page filter, inserts a Timeline linked to that PivotTable, draws a Pie chart that shows only percentages with a 14‑point font, and saves the worksheet (timeline + chart) as a PNG image.
// Keywords: Aspose.Cells | .NET | C# | timeline | pivot table | pie chart | percentage data labels | custom label font | export PNG | chart image generation
// Common Searches: Aspose.Cells add timeline to pivot table and save as PNG | C# pie chart show only percentages Aspose.Cells | change data label font size in Aspose.Cells chart | export worksheet with timeline and chart to image | how to create timeline slicer in Aspose.Cells .NET
// Developer Intent: Generate a worksheet that combines a PivotTable‑driven timeline and a pie chart with percentage‑only labels at a specific font size, then output the result as a PNG file.
// Use Cases: Build an interactive sales dashboard where the timeline filters data and the pie chart visualizes category share with clear percentage labels. | Create printable PNG snapshots of analysis reports that include both a timeline slicer and a chart for slide decks. | Automate email attachments that contain a single image summarizing pivot‑filtered data and its distribution.
// AI Prompts: Write C# code with Aspose.Cells to add a timeline linked to a pivot table and export the sheet as a PNG image. | Show how to configure a pie chart in Aspose.Cells so data labels display only percentages and use a 14‑point font. | Explain how to position a timeline and a chart on the same worksheet before rendering them to an image.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

// Creates a workbook, adds category, value and date data, builds a PivotTable with a date page filter, inserts a Timeline linked to that PivotTable, draws a Pie chart that shows only percentages with a 14‑point font, and saves the worksheet (timeline + chart) as a PNG image.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (used for both the chart and the pivot table)
            cells["A1"].PutValue("Category");
            cells["A2"].PutValue("A");
            cells["A3"].PutValue("B");
            cells["A4"].PutValue("C");

            cells["B1"].PutValue("Value");
            cells["B2"].PutValue(10);
            cells["B3"].PutValue(20);
            cells["B4"].PutValue(30);

            // Add a Date column required for the Timeline (must be a date/time field)
            cells["C1"].PutValue("Date");
            cells["C2"].PutValue(new DateTime(2023, 1, 1));
            cells["C3"].PutValue(new DateTime(2023, 2, 1));
            cells["C4"].PutValue(new DateTime(2023, 3, 1));

            // -------------------------------------------------
            // Create a PivotTable – required as a data source for the Timeline
            // -------------------------------------------------
            PivotTableCollection pivots = worksheet.PivotTables;
            int pivotIndex = pivots.Add("A1:C4", "D1", "PivotTable");
            PivotTable pivot = pivots[pivotIndex];

            // Row field: Category
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            // Data field: Value
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            // Page (filter) field: Date (required for Timeline)
            pivot.AddFieldToArea(PivotFieldType.Page, "Date");

            pivot.RefreshData();
            pivot.CalculateData();

            // -------------------------------------------------
            // Add a Timeline linked to the PivotTable (Date)
            // -------------------------------------------------
            worksheet.Timelines.Add(pivot, 10, 5, "Date");

            // -------------------------------------------------
            // Add a Pie chart and configure data labels
            // -------------------------------------------------
            int chartIndex = worksheet.Charts.Add(ChartType.Pie, 5, 0, 15, 8);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // Enable data labels, show percentages, hide raw values
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowPercentage = true;
            dataLabels.ShowValue = false;

            // Adjust the font size of the data labels
            dataLabels.Font.Size = 14;

            // -------------------------------------------------
            // Save the worksheet (including the Timeline and chart) as a PNG image
            // -------------------------------------------------
            workbook.Save("TimelineChart.png", SaveFormat.Png);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
