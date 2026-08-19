// Title: Set Z‑axis label format to percentage in a 3‑D column chart with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, builds a 3‑D column chart, and formats the Series (Z) axis tick labels as percentages using the number format "0%" before saving the file.
// Keywords: Aspose.Cells | C# chart formatting | Z axis percentage | SeriesAxis tick label format | 3D column chart Aspose | Excel chart number format | Aspose.Cells for .NET | percentage axis label
// Common Searches: Aspose.Cells set Z axis to percent | C# 3D column chart percentage axis | SeriesAxis number format Aspose.Cells | How to display Z axis values as % in Excel chart using Aspose | format chart Z axis percentage Aspose.Cells
// Developer Intent: Apply a percentage number format to the Z‑axis (Series axis) tick labels of a 3‑D column chart in Aspose.Cells for .NET.
// Use Cases: Generate financial or statistical reports where the series axis must show relative values as percentages. | Create interactive Excel dashboards with 3‑D column charts that convey proportionate data clearly. | Automate workbook production that requires custom number formatting on the Z‑axis for downstream analysis.
// AI Prompts: Show C# code to set the Z‑axis tick labels to a percentage format in an Aspose.Cells 3‑D column chart. | How do I apply the number format "0%" to the SeriesAxis of a chart using Aspose.Cells for .NET? | Provide an example of formatting the Z‑axis as percentages and saving the workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, builds a 3‑D column chart, and formats the Series (Z) axis tick labels as percentages using the number format "0%" before saving the file.
    class SetZAxisPercentage
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a 3‑D column chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B2"].PutValue(0.1);
                sheet.Cells["B3"].PutValue(0.2);
                sheet.Cells["B4"].PutValue(0.3);
                sheet.Cells["C2"].PutValue(0.4);
                sheet.Cells["C3"].PutValue(0.5);
                sheet.Cells["C4"].PutValue(0.6);

                // Add a 3‑D column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:C4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories

                // Set Z‑axis (Series axis) tick label format to display percentages
                chart.SeriesAxis.TickLabels.NumberFormat = "0%";

                // Save the workbook
                workbook.Save("ZAxisPercentage.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            SetZAxisPercentage.Run();
        }
    }
}
