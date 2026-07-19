// Title: Aspose.Cells C# – Add calculated data labels to a histogram and auto‑fit label shapes
// Description: This example creates a workbook, fills category and value columns, adds a calculation column (value × 1.1), inserts a histogram chart, links the data labels to the calculated range, positions the labels inside the bars, enables auto‑sizing of each label shape, recalculates the chart, and saves the file as an .xlsx document.
// Keywords: Aspose.Cells histogram | C# chart data labels | link data labels to cell range | auto fit label shape | LinkedSource Aspose.Cells | ShowCellRange | IsResizeShapeToFitText | chart.Calculate | .NET Excel chart example | GitHub Aspose.Cells histogram
// Common Searches: Aspose.Cells add data labels from formula range | C# histogram chart with calculated labels | auto resize chart data label shape Aspose.Cells | link histogram data labels to another column | set data label position inside bars Aspose.Cells | recalculate chart after linking labels
// Developer Intent: Add data labels sourced from a calculated range to a histogram chart and make each label shape automatically fit its text.
// Use Cases: Display adjusted values (e.g., 10 % increase) as labels while keeping the original series unchanged. | Create charts that automatically update label values after formula changes. | Prevent label overlap by auto‑sizing label shapes to accommodate longer numbers or units.
// AI Prompts: Generate C# code using Aspose.Cells to create a histogram chart with data labels linked to a separate calculation column and auto‑fit each label shape to its text. | Show how to set the data label position to InsideBase, enable ShowCellRange, and assign LinkedSource for a histogram in Aspose.Cells. | Explain the steps to recalculate a chart after linking data labels to formula cells with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsHistogramExample
{
    // This example creates a workbook, fills category and value columns, adds a calculation column (value × 1.1), inserts a histogram chart, links the data labels to the calculated range, positions the labels inside the bars, enables auto‑sizing of each label shape, recalculates the chart, and saves the file as an .xlsx document.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Populate raw data (categories) in column A
                // -------------------------------------------------
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                // Populate original values in column B
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(12);
                sheet.Cells["B3"].PutValue(25);
                sheet.Cells["B4"].PutValue(18);

                // -------------------------------------------------
                // Create a calculation range in column C (e.g., value * 1.1)
                // -------------------------------------------------
                sheet.Cells["C1"].PutValue("Calc");
                sheet.Cells["C2"].Formula = "=B2*1.1";
                sheet.Cells["C3"].Formula = "=B3*1.1";
                sheet.Cells["C4"].Formula = "=B4*1.1";

                // -------------------------------------------------
                // Add a histogram chart
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Histogram, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data series to use the original values (column B)
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // -------------------------------------------------
                // Enable data labels for the series
                // -------------------------------------------------
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true; // Show the numeric value
                series.DataLabels.Position = LabelPositionType.InsideBase; // Position inside the bar

                // -------------------------------------------------
                // Link data labels to the calculation range (column C)
                // -------------------------------------------------
                series.DataLabels.ShowCellRange = true; // Enable cell range linking
                series.DataLabels.LinkedSource = "C2:C4"; // Reference the calculated cells

                // -------------------------------------------------
                // Ensure each label shape auto‑fits the text
                // -------------------------------------------------
                foreach (ChartPoint point in series.Points)
                {
                    point.DataLabels.IsResizeShapeToFitText = true; // Auto‑fit shape to text
                    // Optional: customize shape type if desired (removed due to API compatibility)
                }

                // -------------------------------------------------
                // Recalculate the chart so that linked formulas are evaluated
                // -------------------------------------------------
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
