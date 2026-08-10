// Title: C# – Resize Data Label Shapes and Apply Semi‑Transparent Fill to a Stacked Area Chart with Aspose.Cells
// Description: Creates a new workbook, adds a stacked area chart, sets the first series to a 40% transparent blue fill, disables auto‑fit for each point’s data label, assigns a custom width of 60 points, recalculates the chart, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | .NET | stacked area chart | data label size | custom label width | semi transparent fill | disable auto fit | ChartPoint | Excel export | sample code
// Common Searches: Aspose.Cells resize data label shape C# | apply transparency to chart series Aspose.Cells | set fixed width for data labels in stacked area chart | disable auto‑fit for chart data labels .NET | sample code for customizing chart labels Aspose.Cells
// Developer Intent: Adjust the size of data label shapes for each point in a stacked area chart while using a semi‑transparent fill on the series.
// Use Cases: Generate an Excel report where all data labels have a uniform width regardless of text length. | Create visually consistent stacked area charts with partially transparent series fills. | Automate chart styling in server‑side .NET applications that use Aspose.Cells.
// AI Prompts: Write C# code with Aspose.Cells to build a stacked area chart, apply a 40% transparent blue fill to the first series, and set each data label shape width to 60 points while turning off auto‑fit. | Explain how to disable auto‑fit and manually size data label shapes for individual points after changing a series' fill transparency in Aspose.Cells. | Provide step‑by‑step instructions to recalculate a chart and save the workbook after customizing series fill and data label dimensions.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, adds a stacked area chart, sets the first series to a 40% transparent blue fill, disables auto‑fit for each point’s data label, assigns a custom width of 60 points, recalculates the chart, and saves the file as an Excel workbook.
class ResizeDataLabelShapes
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill sample data for the area chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add an area chart (StackedArea may not be available in older versions)
            int chartIndex = sheet.Charts.Add(ChartType.Area, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply semi‑transparent fill to the first series
            Series series = chart.NSeries[0];
            series.Area.FillFormat.SolidFill.Color = Color.Blue;
            series.Area.FillFormat.Transparency = 0.4; // 40% transparent

            // Resize data label shapes for each point in the series
            foreach (ChartPoint point in series.Points)
            {
                // Disable auto‑fit so custom size can be applied
                point.DataLabels.IsResizeShapeToFitText = false;
                // Set a custom width (height adjusts automatically)
                point.DataLabels.Width = 60;
            }

            // Recalculate the chart to apply changes
            chart.Calculate();

            // Save the workbook
            string outputPath = "StackedAreaDataLabelsResized.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
