// Title: C# – Change Chart Value‑Axis Tick‑Label Color by Data Threshold Using Aspose.Cells
// Description: Shows how to build a workbook, insert sample categories and values, apply a three‑color scale conditional format, calculate the maximum cell value, and programmatically set the value‑axis tick‑label font color (and bold style) when the maximum reaches a defined threshold, then save the workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells .NET | C# chart axis formatting | tick label color conditional formatting | value axis font color Aspose.Cells | Excel chart dynamic styling | threshold based axis label | three color scale Aspose.Cells | programmatic Excel chart | conditional formatting chart axis | Aspose.Cells example
// Common Searches: Aspose.Cells change chart axis tick label color | C# set value axis font color based on cell value | conditional formatting affect chart axis labels | how to bold tick labels in Aspose.Cells chart | dynamic chart styling with Aspose.Cells .NET
// Developer Intent: Recolor and optionally bold the chart’s value‑axis tick labels when the data’s maximum value exceeds a specified threshold.
// Use Cases: Highlight high‑risk values in financial dashboards by turning axis labels red when any metric passes a risk limit. | Create self‑updating reports where axis label styling reflects real‑time data thresholds without manual edits. | Emphasize top‑performing categories in sales charts by applying a color‑scale to source cells and changing axis label appearance when sales exceed a target.
// AI Prompts: Generate C# code with Aspose.Cells that sets the value‑axis tick‑label font to orange and bold when the maximum cell value in a range is greater than 100. | Explain how to apply a two‑color conditional format to a range and use the result to conditionally format chart axis labels in Aspose.Cells for .NET. | Provide steps to retrieve the maximum numeric value from a worksheet range and change chart axis label properties based on that value using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Shows how to build a workbook, insert sample categories and values, apply a three‑color scale conditional format, calculate the maximum cell value, and programmatically set the value‑axis tick‑label font color (and bold style) when the maximum reaches a defined threshold, then save the workbook with Aspose.Cells for .NET.
    public class TickLabelConditionalFormattingDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (categories in column A, values in column B)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Low");
                sheet.Cells["B2"].PutValue(15);
                sheet.Cells["A3"].PutValue("Medium");
                sheet.Cells["B3"].PutValue(45);
                sheet.Cells["A4"].PutValue("High");
                sheet.Cells["B4"].PutValue(85);
                sheet.Cells["A5"].PutValue("Very High");
                sheet.Cells["B5"].PutValue(120);

                // Add a column chart that uses the data range
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B5", true);          // Values
                chart.NSeries.CategoryData = "A2:A5";     // Categories

                // Apply a color scale conditional formatting to the value cells (B2:B5)
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

                // Define the area to which the formatting will be applied
                CellArea area = new CellArea { StartRow = 1, EndRow = 4, StartColumn = 1, EndColumn = 1 };
                cfCollection.AddArea(area);

                // Add a 3‑color scale condition
                int conditionIdx = cfCollection.AddCondition(FormatConditionType.ColorScale);
                FormatCondition condition = cfCollection[conditionIdx];
                condition.ColorScale.Is3ColorScale = true;

                // Minimum (green)
                condition.ColorScale.MinCfvo.Type = FormatConditionValueType.Min;
                condition.ColorScale.MinColor = Color.LightGreen;

                // Midpoint (yellow) at 50 % percentile
                condition.ColorScale.MidCfvo.Type = FormatConditionValueType.Percentile;
                condition.ColorScale.MidCfvo.Value = 50;
                condition.ColorScale.MidColor = Color.Yellow;

                // Maximum (red)
                condition.ColorScale.MaxCfvo.Type = FormatConditionValueType.Max;
                condition.ColorScale.MaxColor = Color.IndianRed;

                // Define a threshold for emphasis (e.g., values >= 80)
                double emphasisThreshold = 80;

                // Determine the maximum value in the range to decide the tick‑label color
                double maxValue = double.MinValue;
                for (int row = 1; row <= 4; row++)
                {
                    double val = sheet.Cells[row, 1].DoubleValue;
                    if (val > maxValue) maxValue = val;
                }

                // If the maximum exceeds the threshold, highlight the value‑axis tick labels in red;
                // otherwise keep the default black color.
                TickLabels valueAxisTickLabels = chart.ValueAxis.TickLabels;
                if (maxValue >= emphasisThreshold)
                {
                    valueAxisTickLabels.Font.Color = Color.Red;          // Emphasis color
                    valueAxisTickLabels.Font.IsBold = true;             // Optional visual cue
                }
                else
                {
                    valueAxisTickLabels.Font.Color = Color.Black;
                    valueAxisTickLabels.Font.IsBold = false;
                }

                // Optionally format the tick‑label numbers to show no decimal places
                valueAxisTickLabels.NumberFormat = "0";

                // Save the workbook
                workbook.Save("TickLabelConditionalFormattingDemo.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            TickLabelConditionalFormattingDemo.Run();
        }
    }
}
