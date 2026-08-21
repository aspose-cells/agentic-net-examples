// Title: Create a Pivot Timeline with a Pie Chart, percentage data labels, custom font size, and export to PNG using AspNet.Cells for .NET
// Description: This example builds a workbook, adds a pivot table, links a timeline to the Date field, inserts a pie chart that shows percentages in a 14‑point dark‑blue font, and saves the first worksheet as a PNG image.
// Keywords: Aspose.Cells | C# | .NET | timeline | pivot table | pie chart | percentage data labels | custom label font | PNG export | worksheet image rendering
// Common Searches: Aspose.Cells add timeline to pivot table | C# pie chart show percentages Aspose.Cells | export worksheet as PNG Aspose.Cells .NET | change data label font size Aspose.Cells chart | render timeline with chart to image
// Developer Intent: Render a worksheet that combines a pivot‑driven timeline and a pie chart with styled percentage labels, then save it as a PNG file.
// Use Cases: Build interactive dashboards where a timeline filters a pie chart and the view is captured as an image for reports. | Automate monthly reporting by generating a PNG snapshot of a worksheet that highlights category shares with clear, large labels. | Create presentation‑ready graphics from Excel‑like data without opening Excel, using Aspose.Cells to style and export charts.
// AI Prompts: Generate C# code with Aspose.Cells to attach a timeline to a pivot table, add a pie chart displaying percentage labels in 14‑point dark blue, and save the sheet as PNG. | Explain how to format date cells, customize chart data label appearance, and render only the first worksheet to an image using Aspose.Cells for .NET. | Show the steps to refresh a pivot table after data changes, ensure the timeline reflects the new dates, and export the result as a high‑resolution PNG.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineDemo
{
    // This example builds a workbook, adds a pivot table, links a timeline to the Date field, inserts a pie chart that shows percentages in a 14‑point dark‑blue font, and saves the first worksheet as a PNG image.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data: Category, Date, Value
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Date");
            cells["C1"].PutValue("Value");

            cells["A2"].PutValue("A");
            cells["B2"].PutValue(new DateTime(2023, 1, 1));
            cells["C2"].PutValue(120);

            cells["A3"].PutValue("B");
            cells["B3"].PutValue(new DateTime(2023, 2, 1));
            cells["C3"].PutValue(150);

            cells["A4"].PutValue("C");
            cells["B4"].PutValue(new DateTime(2023, 3, 1));
            cells["C4"].PutValue(90);

            cells["A5"].PutValue("D");
            cells["B5"].PutValue(new DateTime(2023, 4, 1));
            cells["C5"].PutValue(60);

            // Apply date style to the Date column
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "m/d/yyyy";
            for (int row = 2; row <= 5; row++)
            {
                cells[row - 1, 1].SetStyle(dateStyle);
            }

            // Create a PivotTable based on the data
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E1", "SamplePivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Column, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline linked to the PivotTable (using the Date field)
            sheet.Timelines.Add(pivot, "G1", "Date");

            // Add a Pie chart to visualize the same values
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 12, 0, 28, 15);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("C2:C5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Enable data labels, show percentages, and adjust font size
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowPercentage = true;
            dataLabels.ShowValue = false;
            dataLabels.Font.Size = 14;          // Adjust label font size
            dataLabels.Font.Color = Color.DarkBlue;

            // Save the workbook as a PNG image (first worksheet rendered)
            workbook.Save("TimelineWithDataLabels.png", SaveFormat.Png);
        }
    }
}
