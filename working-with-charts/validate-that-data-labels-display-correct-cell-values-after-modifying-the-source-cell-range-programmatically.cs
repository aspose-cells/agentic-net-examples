// Title: Validate Updated Linked Data Labels in an Aspose.Cells .NET Chart (C#)
// Description: C# example that creates a column chart, links its data labels to cells C2:C3, changes those cells, calls chart.Calculate() to refresh the chart, and then verifies each label's Text matches the updated cell value.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# chart data labels | linked data labels | chart.Calculate() | update chart labels programmatically | validate chart labels | Excel chart automation | Aspose.Cells example | data label verification
// Common Searches: how to refresh Aspose.Cells chart after changing linked cells | C# verify chart data labels match updated cell range | Aspose.Cells recalculate chart linked data labels | validate Aspose.Cells chart labels programmatically | chart.Calculate() linked data labels C#
// Developer Intent: Confirm that a chart’s linked data labels reflect the latest cell values after the source range is modified.
// Use Cases: Create a column chart and link its data labels to a specific cell range. | Modify the linked cells and invoke chart.Calculate() to force a refresh. | Iterate through series points and compare DataLabels.Text with the corresponding cell StringValue to assert correctness.
// AI Prompts: Generate C# code using Aspose.Cells to link chart data labels to a cell range, update the cells, recalculate the chart, and validate the label text. | Explain the role of chart.Calculate() in updating linked data labels in Aspose.Cells and show how to compare DataLabels.Text with source cells. | Provide a unit‑test snippet that asserts each Aspose.Cells chart data label equals its linked cell after the cell values are changed.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsDataLabelValidation
{
    // C# example that creates a column chart, links its data labels to cells C2:C3, changes those cells, calls chart.Calculate() to refresh the chart, and then verifies each label's Text matches the updated cell value.
    class Program
    {
        static void Main()
        {
            // ---------- Create workbook and worksheet ----------
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // ---------- Populate source data ----------
            // Category column
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");

            // Value column (used for the chart series)
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["B3"].PutValue(200);

            // Linked source column (used for data labels)
            worksheet.Cells["C1"].PutValue("Label");
            worksheet.Cells["C2"].PutValue("100 units");
            worksheet.Cells["C3"].PutValue("200 units");

            // ---------- Create a column chart ----------
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set chart data range and category data
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // ---------- Configure data labels ----------
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;               // Show the series value
            series.DataLabels.ShowCellRange = true;           // Enable showing linked cell range
            series.DataLabels.LinkedSource = "C2:C3";         // Link to cells C2:C3
            series.DataLabels.Font.Color = Color.Blue;       // Optional styling

            // Save the initial workbook (optional, for inspection)
            workbook.Save("InitialChart.xlsx");

            // ---------- Modify the linked source cells ----------
            worksheet.Cells["C2"].PutValue("150 units");
            worksheet.Cells["C3"].PutValue("250 units");

            // Recalculate the chart so that it picks up the changed cell values
            chart.Calculate();

            // ---------- Validate that data labels reflect the updated cell values ----------
            bool allLabelsMatch = true;
            for (int i = 0; i < series.Points.Count; i++)
            {
                // The text of each data label should now be the updated linked cell value
                string expected = worksheet.Cells[$"C{2 + i}"].StringValue;
                string actual = series.Points[i].DataLabels.Text;

                Console.WriteLine($"Point {i + 1}: Expected = \"{expected}\", Actual = \"{actual}\"");

                if (!string.Equals(expected, actual, StringComparison.Ordinal))
                {
                    allLabelsMatch = false;
                }
            }

            Console.WriteLine(allLabelsMatch
                ? "All data labels display the correct updated cell values."
                : "Data label validation failed.");

            // Save the final workbook with updated labels
            workbook.Save("UpdatedChart.xlsx");
        }
    }
}
