// Title: C# Aspose.Cells – Bind a Chart to Merged Cells Using SetChartDataRange
// Description: Demonstrates how to create a workbook, merge category cells, add a column chart, and bind the chart to a range that contains merged cells (A1:B7) with SetChartDataRange so each merged area is treated as a single data point. The workbook is saved as MergedCellsChartOutput.xlsx.
// Keywords: Aspose.Cells SetChartDataRange | C# chart merged cells | bind chart to merged range .NET | column chart merged categories | Aspose.Cells merged cell handling | SetChartDataRange example C# | Aspose.Cells chart data range
// Common Searches: Aspose.Cells bind chart to merged cells C# | SetChartDataRange merged cells example | Create column chart with merged row headers Aspose | How to handle merged cells in Aspose.Cells charts | C# Aspose.Cells chart series from merged range
// Developer Intent: The developer needs to link a chart to a data range that includes merged cells, ensuring the chart interprets each merged block as one category value.
// Use Cases: Produce a column chart where category labels span multiple rows by merging cells and using SetChartDataRange for automatic series creation. | Generate financial or survey reports with grouped rows under merged headers and visualize the groups in a single chart without manual series definition. | Automate workbook creation for dashboards where merged row titles must be reflected correctly in chart legends.
// AI Prompts: Show C# code that binds an Aspose.Cells chart to a range containing merged cells using SetChartDataRange. | Explain how SetChartDataRange processes merged cells when the by‑column flag is true in Aspose.Cells for .NET. | Provide a step‑by‑step guide to create a column chart from merged category cells with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsMergedChartDemo
{
    // Demonstrates how to create a workbook, merge category cells, add a column chart, and bind the chart to a range that contains merged cells (A1:B7) with SetChartDataRange so each merged area is treated as a single data point. The workbook is saved as MergedCellsChartOutput.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Prepare data for the chart
            // ------------------------------------------------------------
            // Header row
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");

            // Row 2 – Category "Group 1" will span two rows (merged)
            cells["A2"].PutValue("Group 1");
            cells["B2"].PutValue(40);   // First value for Group 1
            // Row 3 – second part of the merged category (value only)
            cells["B3"].PutValue(60);   // Second value for Group 1

            // Row 4 – Category "Group 2" (no merge)
            cells["A4"].PutValue("Group 2");
            cells["B4"].PutValue(80);

            // Row 5 – Category "Group 3" will span three rows (merged)
            cells["A5"].PutValue("Group 3");
            cells["B5"].PutValue(30);
            cells["B6"].PutValue(50);
            cells["B7"].PutValue(70);

            // Merge the category cells for Group 1 (A2:A3) and Group 3 (A5:A7)
            // The merged cell's value is taken from the upper‑left cell (A2 and A5)
            cells.Merge(1, 0, 2, 1); // A2:A3
            cells.Merge(4, 0, 3, 1); // A5:A7

            // ------------------------------------------------------------
            // Add a column chart and bind it to the data range that includes merged cells
            // ------------------------------------------------------------
            // The data range includes the header row and all data rows.
            // SetChartDataRange will treat each merged cell as a single data point,
            // using the value from the top‑left cell of the merged area.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 9, 0, 25, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Bind the chart to the range A1:B7 (including merged cells)
            // 'true' indicates that the series are plotted by column.
            chart.SetChartDataRange("A1:B7", true);

            // Optional: set a title for clarity
            chart.Title.Text = "Values by Category (Merged Cells)";

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            workbook.Save("MergedCellsChartOutput.xlsx");
        }
    }
}
