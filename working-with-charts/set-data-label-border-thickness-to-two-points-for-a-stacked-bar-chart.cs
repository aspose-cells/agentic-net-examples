// Title: Aspose.Cells .NET – Set Data Label Border Thickness to 2 pt in a Stacked Bar Chart
// Description: C# example that creates a workbook, adds sample data, inserts a stacked bar chart, enables data labels for the first series, makes the label border visible, sets its weight to 2 points, applies a black color, and saves the file as XLSX.
// Keywords: Aspose.Cells data label border | stacked bar chart label thickness | C# chart formatting Aspose.Cells | set label border weight 2 points | Excel chart data label border color | Aspose.Cells chart customization | chart series label border visibility
// Common Searches: Aspose.Cells set data label border thickness .NET | C# stacked bar chart label border weight | how to change chart data label border in Aspose.Cells | Aspose.Cells label border 2 pt example | make data label border visible in Excel chart using Aspose
// Developer Intent: Apply a 2‑point black border to the data labels of the first series in a stacked bar chart using Aspose.Cells for .NET.
// Use Cases: Highlight values in a stacked bar chart with a thicker border for better visual emphasis. | Standardize label appearance across reports by setting consistent border weight and color. | Create Excel dashboards where data labels need a distinct outline to improve readability.
// AI Prompts: Generate C# code with Aspose.Cells that adds a stacked bar chart and sets the first series data label border to 2 pt black. | Show how to modify data label border visibility, weight, and color for multiple series in an Aspose.Cells chart. | Explain step‑by‑step how to customize chart data label borders (visibility, thickness, color) in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, adds sample data, inserts a stacked bar chart, enables data labels for the first series, makes the label border visible, sets its weight to 2 points, applies a black color, and saves the file as XLSX.
    public class StackedBarDataLabelBorderDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a stacked bar chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Add a stacked bar chart (use ChartType.BarStacked)
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add the two series to the chart
                chart.NSeries.Add("B2:B4", true); // Series1 values
                chart.NSeries.Add("C2:C4", true); // Series2 values
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                DataLabels labels = series.DataLabels;
                labels.ShowValue = true; // display the value

                // Set the border thickness of the data labels to 2 points
                labels.Border.IsVisible = true;
                labels.Border.WeightPt = 2.0; // two points thickness
                labels.Border.Color = Color.Black; // optional: set border color

                // Save the workbook
                string outputPath = "StackedBarDataLabelBorderDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            StackedBarDataLabelBorderDemo.Run();
        }
    }
}
