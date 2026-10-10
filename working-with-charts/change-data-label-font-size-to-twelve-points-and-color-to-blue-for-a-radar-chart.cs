// Title: Set radar chart data label font size to 12 points and color to blue with Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, add sample data, generate a radar chart, and apply a 12‑pt blue font to all data labels using Aspose.Cells in C#. | Modify an existing radar chart so that its series data labels display values in a twelve‑point blue typeface with Aspose.Cells for .NET. | Write C# code that builds a radar chart and formats the data labels to use a 12‑point blue font.
// Common Searches: Aspose.Cells how to change radar chart data label font size to 12 points | C# set data label color to blue in Excel radar chart using Aspose.Cells | formatting data labels on a radar chart with Aspose.Cells .NET example | sample code for customizing radar chart labels font and color in Aspose.Cells
// Tags: Aspose.Cells radar chart data label formatting | C# set data label font size Aspose.Cells | Aspose.Cells data label color blue | Excel radar chart label styling with Aspose | Aspose.Cells chart series label customization

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills sample data, adds a radar chart, enables data labels, and sets their font to 12 pt blue before saving as RadarChartWithLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data for the radar chart
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Series1");
            ws.Cells["C1"].PutValue("Series2");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["A3"].PutValue("B");
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["B4"].PutValue(30);
            ws.Cells["C2"].PutValue(15);
            ws.Cells["C3"].PutValue(25);
            ws.Cells["C4"].PutValue(35);

            // Add a radar chart to the worksheet
            int chartIndex = ws.Charts.Add(ChartType.Radar, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIndex];

            // Set the data range for the series (including categories)
            chart.NSeries.Add("B2:C4", true);

            // Show data labels for each series
            foreach (Series series in chart.NSeries)
            {
                series.DataLabels.ShowValue = true;               // display values
                series.DataLabels.Font.Size = 12;                 // set font size
                series.DataLabels.Font.Color = Color.Blue;       // set font color
            }

            // Define output file name
            string outputPath = "RadarChartWithLabels.xlsx";

            // Save the workbook with the radar chart
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
