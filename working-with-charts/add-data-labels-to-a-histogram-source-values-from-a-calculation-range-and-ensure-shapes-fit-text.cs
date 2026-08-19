// Title: Add Calculated Data Labels to a Histogram Chart with Auto‑Fit Shapes – Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills columns A‑B with categories and values, builds a formula range in column C, adds a column (histogram‑style) chart, links data labels to the formula range, enables shape auto‑sizing, sets label position to InsideBase, recalculates the chart, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells histogram data labels | C# chart linked source labels | auto resize label shape Aspose.Cells | custom chart labels from formulas | SetLabelPosition InsideBase Aspose.Cells | Excel chart data labels .NET
// Common Searches: link chart data labels to a cell range Aspose.Cells | auto‑fit data label shape in Excel using C# | histogram with concatenated labels Aspose.Cells | how to use LinkedSource for chart labels .NET | set data label position inside base column chart
// Developer Intent: Generate a histogram where each column displays a calculated label that automatically resizes to fit its text.
// Use Cases: Display "Category: Value" labels on a column chart without manual text entry. | Produce Excel reports where label boxes expand to accommodate dynamic formula results. | Create a reusable routine that adds formula‑driven, auto‑sized labels to any chart series.
// AI Prompts: Write C# code with Aspose.Cells that adds data labels to a column chart, pulls label text from a calculated range, and enables auto‑fit of label shapes. | Show how to link a histogram's data labels to concatenated formulas and set the label position to InsideBase using Aspose.Cells for .NET. | Explain the steps to recalculate a chart after assigning a LinkedSource to data labels in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsHistogramDataLabels
{
    // Creates a workbook, fills columns A‑B with categories and values, builds a formula range in column C, adds a column (histogram‑style) chart, links data labels to the formula range, enables shape auto‑sizing, sets label position to InsideBase, recalculates the chart, and saves the file as an Excel workbook.
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
                // Populate source data for the histogram (Category and Value)
                // -------------------------------------------------
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["A5"].PutValue("D");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(15);
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["B4"].PutValue(45);
                sheet.Cells["B5"].PutValue(60);

                // -------------------------------------------------
                // Create a calculation range that derives a label text.
                // For demonstration, concatenate the category with the value.
                // The result will be placed in column C.
                // -------------------------------------------------
                sheet.Cells["C1"].PutValue("Label");
                sheet.Cells["C2"].Formula = "=A2 & \": \" & B2";
                sheet.Cells["C3"].Formula = "=A3 & \": \" & B3";
                sheet.Cells["C4"].Formula = "=A4 & \": \" & B4";
                sheet.Cells["C5"].Formula = "=A5 & \": \" & B5";

                // -------------------------------------------------
                // Add a histogram‑like chart (using Column) to the worksheet
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series (values only)
                chart.NSeries.Add("B2:B5", true);
                // Set the category (X‑axis) data
                chart.NSeries.CategoryData = "A2:A5";

                // -------------------------------------------------
                // Enable data labels for the series
                // -------------------------------------------------
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;                     // Show the numeric value
                series.DataLabels.ShowCellRange = true;                // Use a cell range as the source for label text
                series.DataLabels.LinkedSource = "C2:C5";               // Link to the calculation range created above
                series.DataLabels.IsResizeShapeToFitText = true;       // Ensure the label shape auto‑fits the text

                // Optional: set a more compact position so labels do not overlap
                series.DataLabels.Position = LabelPositionType.InsideBase;

                // -------------------------------------------------
                // Recalculate the chart to apply the linked source values
                // -------------------------------------------------
                chart.Calculate();

                // -------------------------------------------------
                // Save the workbook
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
