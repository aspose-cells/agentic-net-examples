using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsDataLabelValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // ---------- Populate initial data ----------
            // Category column
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");

            // Value column (numeric data for the chart)
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["B3"].PutValue(200);

            // Linked source column for data labels (text to be shown)
            worksheet.Cells["C1"].PutValue("Label");
            worksheet.Cells["C2"].PutValue("100 units");
            worksheet.Cells["C3"].PutValue("200 units");

            // ---------- Create a column chart ----------
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the chart data range and category data
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // ---------- Configure data labels ----------
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;          // Show the numeric value
            series.DataLabels.ShowCellRange = true;      // Enable showing linked cell range
            series.DataLabels.LinkedSource = "C2:C3";    // Link to the label cells
            series.DataLabels.Font.Color = Color.Blue;   // Optional styling

            // Save the workbook after initial setup (optional)
            workbook.Save("ChartWithDataLabels_Initial.xlsx");

            // ---------- Modify the linked source cells ----------
            worksheet.Cells["C2"].PutValue("150 units");
            worksheet.Cells["C3"].PutValue("250 units");

            // Recalculate the chart to reflect the changes
            chart.Calculate();

            // ---------- Validate that data labels display the updated cell values ----------
            bool allLabelsMatch = true;
            for (int i = 0; i < series.Points.Count; i++)
            {
                // Expected text from the linked source cells
                string expected = worksheet.Cells[$"C{i + 2}"].StringValue;

                // Actual text displayed on the data label
                string actual = series.Points[i].DataLabels.Text;

                Console.WriteLine($"Point {i + 1}: Expected = \"{expected}\", Actual = \"{actual}\"");

                if (!string.Equals(expected, actual, StringComparison.Ordinal))
                {
                    allLabelsMatch = false;
                }
            }

            Console.WriteLine(allLabelsMatch
                ? "All data labels correctly reflect the updated cell values."
                : "Data label validation failed.");

            // Save the final workbook
            workbook.Save("ChartWithDataLabels_Updated.xlsx");
        }
    }
}