// Title: Add calculated data labels to a histogram chart and auto‑fit label shapes with Aspose.Cells for .NET
// AI Prompts: Create a column histogram in Aspose.Cells, bind its data labels to a separate calculation range, and turn on shape auto‑sizing so each label fits its text. | Programmatically set custom text for each point’s data label from a formula cell and ensure the label shape resizes automatically in C#. | Recalculate the chart after linking data labels to formula cells to display calculated values on a histogram using Aspose.Cells.
// Common Searches: how to link histogram data labels to a formula range using Aspose.Cells C# | auto resize chart data label shapes to fit text in Aspose.Cells .NET | display calculated values as data labels on an Excel column chart with Aspose.Cells | customize individual point labels in a histogram chart programmatically Aspose.Cells | set linked source for data labels in Aspose.Cells chart
// Tags: histogram data labels linked to formula range Aspose.Cells | auto‑fit chart label shape C# Aspose.Cells | custom point data label text Aspose.Cells | column chart series data label customization .NET | link data label to cell range Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsHistogramExample
{
    // The example creates a workbook, fills columns with categories, original values, and a calculated range (doubling the values). It adds a column histogram, assigns the value series and categories, enables data labels, links those labels to the calculated range, and forces each label shape to auto‑fit its text. It also demonstrates overriding each point’s label with custom text from the calculation cells, recalculates the chart, and saves the file as HistogramWithCalculatedDataLabels.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate source data (Category and Values)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Create a calculation range (e.g., double the original values)
                sheet.Cells["C1"].PutValue("CalcValue");
                sheet.Cells["C2"].Formula = "=B2*2";
                sheet.Cells["C3"].Formula = "=B3*2";
                sheet.Cells["C4"].Formula = "=B4*2";

                // Add a histogram (using Column chart type)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories

                // Enable data labels for the series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;        // Show the original value
                series.DataLabels.Position = LabelPositionType.InsideBase;

                // Use the calculation range as the source for data label text
                series.DataLabels.ShowCellRange = true;    // Enable cell range display
                series.DataLabels.LinkedSource = "C2:C4";  // Link to calculated values

                // Ensure each data label shape auto‑fits the text
                series.DataLabels.IsResizeShapeToFitText = true;

                // Optionally, customize each point's label (demonstrates Text property)
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];
                    // Override auto‑text with custom text if needed
                    point.DataLabels.IsAutoText = false;
                    string cellRef = $"C{i + 2}";
                    point.DataLabels.Text = $"Calc: {sheet.Cells[cellRef].StringValue}";
                    // Ensure shape fits the custom text
                    point.DataLabels.IsResizeShapeToFitText = true;
                }

                // Recalculate the chart to apply linked formulas and layout
                chart.Calculate();

                // Save the workbook
                workbook.Save("HistogramWithCalculatedDataLabels.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
