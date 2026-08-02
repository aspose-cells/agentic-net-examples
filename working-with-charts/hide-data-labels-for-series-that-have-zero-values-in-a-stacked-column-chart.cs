// Title: Hide Zero-Value Data Labels in a Stacked Column Chart – Aspose.Cells for .NET
// Description: Creates a workbook with sample data that includes zeros, adds a stacked column chart, turns on data labels, then iterates each series point, checks the cell value and disables the label when the value is zero, finally saving the file as XLSX.
// Keywords: Aspose.Cells | C# | .NET chart API | stacked column chart | hide zero data labels | chart point label control | Excel automation | data label visibility | chart customization | zero-value suppression
// Common Searches: Aspose.Cells hide zero data label stacked column | C# remove label for zero value points in Excel chart | how to suppress zero-value labels using Aspose.Cells | programmatically hide specific chart point labels .NET | Aspose.Cells chart label visibility for empty cells
// Developer Intent: Programmatically suppress data labels for points with a value of zero in a stacked column chart using Aspose.Cells for .NET.
// Use Cases: Sales dashboards where categories with no sales should not display a label, keeping the chart clean. | Financial reports that omit zero‑balance entries from stacked column visualizations to avoid clutter. | Automated KPI sheets that automatically hide labels for metrics with no activity during a period.
// AI Prompts: Generate C# code with Aspose.Cells that hides data labels for zero-valued points in any chart type. | Explain how to extend the example to also hide labels for negative values in a stacked column chart. | Suggest a more efficient method to suppress zero-value labels without manually parsing cell addresses.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook with sample data that includes zeros, adds a stacked column chart, turns on data labels, then iterates each series point, checks the cell value and disables the label when the value is zero, finally saving the file as XLSX.
class HideZeroDataLabels
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (including zero values)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");

            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(0);

            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(0);
            sheet.Cells["C3"].PutValue(20);

            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["C4"].PutValue(5);

            // Add a stacked column chart
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add the two series to the chart
            chart.NSeries.Add("B2:B4", true); // Series1
            chart.NSeries.Add("C2:C4", true); // Series2
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for all series
            foreach (Series series in chart.NSeries)
            {
                series.DataLabels.ShowValue = true;
            }

            // Hide data labels for points whose value is zero
            foreach (Series series in chart.NSeries)
            {
                // Get the source range of the series values (e.g., "B2:B4")
                string valuesRange = series.Values;
                if (string.IsNullOrEmpty(valuesRange))
                    continue;

                // Split the range into start and end cell names
                string[] parts = valuesRange.Split(':');
                if (parts.Length != 2)
                    continue;

                // Convert cell names to row/column indices
                int startRow, startColumn, endRow, endColumn;
                CellsHelper.CellNameToIndex(parts[0], out startRow, out startColumn);
                CellsHelper.CellNameToIndex(parts[1], out endRow, out endColumn);

                // Iterate through each point in the series
                for (int i = 0; i < series.Points.Count; i++)
                {
                    // Calculate the corresponding cell coordinates
                    int row = startRow + i;
                    int col = startColumn;

                    // Retrieve the numeric value from the worksheet
                    double pointValue = sheet.Cells[row, col].DoubleValue;

                    // If the value is zero, hide its data label
                    if (Math.Abs(pointValue) < double.Epsilon)
                    {
                        series.Points[i].DataLabels.ShowValue = false;
                    }
                }
            }

            // Save the workbook
            workbook.Save("StackedColumn_HideZeroLabels.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
