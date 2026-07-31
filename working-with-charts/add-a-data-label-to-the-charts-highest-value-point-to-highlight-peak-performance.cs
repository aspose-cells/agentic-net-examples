// Title: Add a Peak Data Label to the Highest Value Point in an Aspose.Cells Column Chart (C#)
// Description: Creates a workbook with month‑sales data, builds a column chart, finds the series point with the maximum Y value, and attaches a custom data label that shows the value, displays the text "Peak", and positions the label above the column before saving the file.
// Keywords: Aspose.Cells C# chart data label | add label to max point Aspose.Cells | highlight peak value column chart | custom chart point label Aspose | C# find maximum chart series value
// Common Searches: Aspose.Cells add data label to highest chart point C# | C# label max value in column chart Aspose.Cells | how to show custom text on chart point Aspose.Cells | position chart data label above column Aspose | find and label peak value in Aspose.Cells chart
// Developer Intent: Add a data label to the chart's highest value point to highlight peak performance.
// Use Cases: Mark the month with the highest sales in a monthly sales column chart. | Emphasize the maximum reading in a scientific measurement chart for quick visual analysis. | Identify the top scorer in a leaderboard chart with a custom "Peak" label.
// AI Prompts: Generate C# code using Aspose.Cells that locates the maximum Y value in a column chart series and adds a data label with the text "Peak" positioned above the column. | Show an Aspose.Cells example that iterates over chart points, finds the highest value, and sets ShowValue and custom Text for its DataLabels. | Explain how to change the label position, font, and formatting for the peak point in an Aspose.Cells column chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook with month‑sales data, builds a column chart, finds the series point with the maximum Y value, and attaches a custom data label that shows the value, displays the text "Peak", and positions the label above the column before saving the file.
    class AddPeakDataLabel
    {
        static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (Month vs Sales)
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");   sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Feb");   sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue("Mar");   sheet.Cells["B4"].PutValue(150);
            sheet.Cells["A5"].PutValue("Apr");   sheet.Cells["B5"].PutValue(250); // highest value
            sheet.Cells["A6"].PutValue("May");   sheet.Cells["B6"].PutValue(180);

            // Add a column chart covering the data range
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B6", true);          // values
            chart.NSeries.CategoryData = "A2:A6";      // categories

            // Locate the point with the maximum Y value in the first series
            Series series = chart.NSeries[0];
            double maxValue = double.MinValue;
            int maxPointIndex = -1;

            for (int i = 0; i < series.Points.Count; i++)
            {
                // YValue is returned as object; convert to double safely
                double y = Convert.ToDouble(series.Points[i].YValue);
                if (y > maxValue)
                {
                    maxValue = y;
                    maxPointIndex = i;
                }
            }

            // If a maximum point was found, add a custom data label to it
            if (maxPointIndex >= 0)
            {
                ChartPoint peakPoint = series.Points[maxPointIndex];
                peakPoint.DataLabels.ShowValue = true;                     // show the numeric value
                peakPoint.DataLabels.Position = LabelPositionType.Above;   // place label above the column
                peakPoint.DataLabels.Text = "Peak";                        // custom text indicating peak performance
            }

            // Save the workbook with the chart and highlighted peak label
            string outputPath = "ChartPeakLabel.xlsx";
            workbook.Save(outputPath);
        }
    }
}
