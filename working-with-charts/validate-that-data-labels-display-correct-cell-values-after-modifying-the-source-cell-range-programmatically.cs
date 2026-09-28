// Title: How to verify that Aspose.Cells .NET chart data labels update after changing linked cells programmatically
// AI Prompts: Generate C# code that links a column chart's data labels to a specific cell range, modifies those cells, calls chart.Calculate, and checks IsChartDataChanged before and after the calculation. | Write a verification loop that iterates through each series point, reads the DataLabels.Text, and compares it to the corresponding worksheet cell value to confirm the label reflects the updated content.
// Common Searches: Aspose.Cells .NET check if chart data changed after editing linked label cells | C# update Excel chart data labels from cell range and validate with IsChartDataChanged | How to programmatically refresh Aspose.Cells chart labels after modifying source cells
// Tags: link chart data labels to cell range Aspose.Cells | chart.IsChartDataChanged usage Aspose.Cells | chart.Calculate refresh data labels Aspose.Cells | validate data label text against worksheet cells C# | Aspose.Cells column chart programmatic update

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsDataLabelValidation
{
    // Creates a workbook, adds a column chart, links its data labels to cells C2:C3, updates those cells, uses IsChartDataChanged to detect pending changes, calls chart.Calculate to refresh, then iterates through series points to confirm each label text matches the updated cell value, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Populate source data for the chart
            // -------------------------------------------------
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["B3"].PutValue(200);

            // Cells that will be linked to data labels
            worksheet.Cells["C2"].PutValue("100 units");
            worksheet.Cells["C3"].PutValue("200 units");

            // -------------------------------------------------
            // 2. Create a column chart
            // -------------------------------------------------
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // -------------------------------------------------
            // 3. Configure data labels to show linked cell range
            // -------------------------------------------------
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;               // Show the numeric value (optional)
            series.DataLabels.ShowCellRange = true;           // Enable showing linked cell range
            series.DataLabels.LinkedSource = "C2:C3";         // Link to custom label cells
            series.DataLabels.Font.Color = Color.Blue;        // Visual cue

            // Save the workbook after initial setup (optional)
            workbook.Save("InitialChart.xlsx");

            // -------------------------------------------------
            // 4. Modify the linked source cells programmatically
            // -------------------------------------------------
            worksheet.Cells["C2"].PutValue("150 units");
            worksheet.Cells["C3"].PutValue("250 units");

            // -------------------------------------------------
            // 5. Detect if chart data has changed before recalculation
            // -------------------------------------------------
            bool changedBeforeCalc = chart.IsChartDataChanged(); // Expected: true because source changed
            Console.WriteLine($"Chart data changed before Calculate(): {changedBeforeCalc}");

            // -------------------------------------------------
            // 6. Recalculate the chart to refresh data labels
            // -------------------------------------------------
            chart.Calculate();

            // Verify that the chart now reports no pending changes
            bool changedAfterCalc = chart.IsChartDataChanged(); // Expected: false
            Console.WriteLine($"Chart data changed after Calculate(): {changedAfterCalc}");

            // -------------------------------------------------
            // 7. Validate that data labels display the updated cell values
            // -------------------------------------------------
            bool allLabelsMatch = true;
            for (int i = 0; i < series.Points.Count; i++)
            {
                // The Text property reflects the linked cell content when ShowCellRange is true
                string labelText = series.Points[i].DataLabels.Text;
                string expectedText = worksheet.Cells[$"C{i + 2}"].StringValue; // C2, C3, ...

                Console.WriteLine($"Label {i}: '{labelText}' (expected: '{expectedText}')");

                if (!labelText.Equals(expectedText, StringComparison.Ordinal))
                {
                    allLabelsMatch = false;
                }
            }

            Console.WriteLine($"All data labels reflect updated source cells: {allLabelsMatch}");

            // -------------------------------------------------
            // 8. Save the final workbook with updated labels
            // -------------------------------------------------
            workbook.Save("ChartWithUpdatedDataLabels.xlsx");
        }
    }
}
