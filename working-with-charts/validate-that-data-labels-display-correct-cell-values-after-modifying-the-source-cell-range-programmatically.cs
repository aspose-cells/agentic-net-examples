// Title: Validate chart data‑label refresh after programmatic change of its source range – Aspose.Cells for .NET (C#)
// Description: The sample creates a workbook, adds a column chart, ties the labels to cells C2:C3, updates those cells, calls chart.Calculate(), and checks that each label’s Text matches the new content before saving the file.
// Keywords: Aspose.Cells | C# | .NET | chart data labels | source range | chart.Calculate | column chart | Excel automation | label validation | programmatic chart update | unit testing
// Common Searches: Aspose.Cells link data labels to a range | update chart labels after cell change .NET | refresh chart after modifying source range | C# verify chart data label text | Aspose.Cells recalculate chart | automated test chart labels
// Developer Intent: Confirm that a chart’s data‑label text reflects runtime changes made to its linked source range.
// Use Cases: Bind label text to a worksheet range and have it update automatically when the range changes. | Include chart label verification in continuous‑integration pipelines for Excel report generation. | Build dynamic dashboards where label content is driven by calculated cells such as units or percentages. | Produce Excel workbooks with charts that display custom text derived from formulas or external data.
// AI Prompts: Write C# code using Aspose.Cells to link chart data labels to a cell range and refresh them after the cells are updated. | Provide a .NET unit test that asserts DataLabels.Text equals expected strings after modifying the linked range. | Explain how Chart.Calculate() updates linked data‑label values in Aspose.Cells. | Create a step‑by‑step tutorial for linking and validating chart labels in an Excel workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsDataLabelValidation
{
    // The sample creates a workbook, adds a column chart, ties the labels to cells C2:C3, updates those cells, calls chart.Calculate(), and checks that each label’s Text matches the new content before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // ---------- Set up initial data ----------
            // Category labels
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");

            // Numeric values for the series
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["B3"].PutValue(200);

            // Text values that will be linked to data labels
            worksheet.Cells["C2"].PutValue("100 units");
            worksheet.Cells["C3"].PutValue("200 units");

            // ---------- Create a column chart ----------
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Configure data labels to show linked cell values
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;          // Show the numeric value (optional)
            series.DataLabels.ShowCellRange = true;      // Enable showing linked cell range
            series.DataLabels.LinkedSource = "C2:C3";    // Link to cells C2:C3
            series.DataLabels.Font.Color = Color.Blue;  // Visual styling (optional)

            // ---------- Verify initial data label texts ----------
            Console.WriteLine("Initial Data Labels:");
            for (int i = 0; i < series.Points.Count; i++)
            {
                string labelText = series.Points[i].DataLabels.Text;
                Console.WriteLine($" Point {i}: {labelText}");
            }

            // ---------- Modify the linked source cells programmatically ----------
            worksheet.Cells["C2"].PutValue("150 units");
            worksheet.Cells["C3"].PutValue("250 units");

            // Recalculate the chart so it picks up the changed cell values
            chart.Calculate();

            // ---------- Validate that data labels reflect the updated cell values ----------
            Console.WriteLine("\nData Labels After Updating Source Cells:");
            bool allMatch = true;
            string[] expected = { "150 units", "250 units" };
            for (int i = 0; i < series.Points.Count; i++)
            {
                string actualLabel = series.Points[i].DataLabels.Text;
                Console.WriteLine($" Point {i}: {actualLabel}");
                if (actualLabel != expected[i])
                {
                    allMatch = false;
                }
            }

            Console.WriteLine("\nValidation Result: " + (allMatch ? "PASS" : "FAIL"));

            // Save the workbook (the chart image will reflect the updated labels)
            workbook.Save("DataLabelsValidationDemo.xlsx");
        }
    }
}
