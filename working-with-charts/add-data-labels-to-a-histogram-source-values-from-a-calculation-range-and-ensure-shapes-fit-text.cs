// Title: Add Calculated Data Labels to a Histogram and Auto‑Fit Shapes with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, populate categories and values, compute a secondary range (e.g., double the original values), insert a column chart used as a histogram, link data‑label text to the calculated range, enable automatic shape resizing and font scaling, position labels inside the columns, recalculate the chart, and save the file as an Excel workbook.
// Keywords: Aspose.Cells histogram data labels | C# chart data label linked source | auto‑fit label shape Aspose.Cells | IsResizeShapeToFitText | AutoScaleFont Aspose.Cells | LabelPositionType.InsideBase | column chart calculated labels .NET | Excel automation Aspose.Cells | dynamic data labels C# | chart recalculation Aspose.Cells
// Common Searches: link data label to cell range Aspose.Cells | auto size data label shape in column chart .NET | histogram with custom calculated labels C# | set label position inside base Aspose.Cells | Aspose.Cells example for data label formulas
// Developer Intent: Generate a histogram where each bar’s label shows a value derived from a formula range and automatically adjusts its shape to the text.
// Use Cases: Sales distribution report that displays a computed metric (e.g., double sales) as a label on each histogram bar. | Dynamic workbook where label values update automatically when source data changes, with labels fitting the column width. | Dashboard chart with inside‑base labels that scale font size to remain readable across varying bar heights.
// AI Prompts: Write C# code using Aspose.Cells to create a histogram, link its data labels to a calculation column, and enable shape auto‑fit. | Show how to set data label position to InsideBase and turn on AutoScaleFont for a column chart in Aspose.Cells. | Provide an example that updates the formula range for label values and refreshes the chart labels in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsHistogramDataLabels
{
    // Demonstrates how to create a workbook, populate categories and values, compute a secondary range (e.g., double the original values), insert a column chart used as a histogram, link data‑label text to the calculated range, enable automatic shape resizing and font scaling, position labels inside the columns, recalculate the chart, and save the file as an Excel workbook.
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
                // 1. Populate source data (Category in A, Values in B)
                // -------------------------------------------------
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["A5"].PutValue("D");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(40);

                // -------------------------------------------------
                // 2. Add a calculation range (C column) that will be used for data‑label text
                //    Example: double the original value
                // -------------------------------------------------
                sheet.Cells["C1"].PutValue("CalcValue");
                for (int row = 2; row <= 5; row++)
                {
                    // Formula: =B2*2, =B3*2, ...
                    sheet.Cells[$"C{row}"].Formula = $"=B{row}*2";
                }

                // -------------------------------------------------
                // 3. Insert a column chart (used as histogram) and bind data
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Series data (values) and category (labels)
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // -------------------------------------------------
                // 4. Enable data labels and link them to the calculation range
                // -------------------------------------------------
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;                 // Show the value (optional)
                series.DataLabels.ShowCellRange = true;             // Use cell range for label text
                series.DataLabels.LinkedSource = "C2:C5";           // Calculation range
                series.DataLabels.IsResizeShapeToFitText = true;   // Auto‑fit shape to text
                series.DataLabels.AutoScaleFont = true;            // Font scales with shape size

                // Optional: set label position inside the bars
                series.DataLabels.Position = LabelPositionType.InsideBase;

                // -------------------------------------------------
                // 5. Recalculate the chart to apply custom positions/sizes
                // -------------------------------------------------
                chart.Calculate();

                // -------------------------------------------------
                // 6. Save the workbook
                // -------------------------------------------------
                workbook.Save("HistogramWithCalculatedDataLabels.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
