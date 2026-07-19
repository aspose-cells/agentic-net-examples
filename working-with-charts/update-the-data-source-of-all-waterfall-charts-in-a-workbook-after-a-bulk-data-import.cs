// Title: Refresh Waterfall Chart Data Sources After Bulk Import with Aspose.Cells for .NET
// Description: Load a workbook, import new values, call Worksheets.RefreshAll to update pivot‑based Waterfall charts, use Workbook.CalculateFormula for non‑pivot charts, and save the refreshed file.
// Keywords: Aspose.Cells | C# | .NET | Waterfall chart | refresh chart data | bulk data import | Worksheets.RefreshAll | Workbook.CalculateFormula | pivot chart update | programmatic chart refresh
// Common Searches: Aspose.Cells refresh waterfall chart after data import | C# update chart data source programmatically | How to recalculate charts in Aspose.Cells | Refresh all charts in workbook .NET | Update pivot Waterfall chart Aspose.Cells
// Developer Intent: Programmatically update every Waterfall chart’s data source in a workbook after performing a bulk data import.
// Use Cases: Refresh pivot‑based Waterfall charts after bulk filling of a data range. | Force full workbook calculation to synchronize non‑pivot Waterfall charts with new values. | Save the workbook so all Waterfall charts display the latest imported data.
// AI Prompts: Generate C# code using Aspose.Cells that refreshes all Waterfall charts after a bulk data import. | Explain how Worksheets.RefreshAll and Workbook.CalculateFormula affect Waterfall chart data sources. | Show an example of importing data into a worksheet and then updating both pivot and non‑pivot Waterfall charts with Aspose.Cells.

using System;
using Aspose.Cells;

namespace UpdateWaterfallCharts
{
    // Load a workbook, import new values, call Worksheets.RefreshAll to update pivot‑based Waterfall charts, use Workbook.CalculateFormula for non‑pivot charts, and save the refreshed file.
    class Program
    {
        static void Main()
        {
            // Load the workbook that contains Waterfall charts.
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // -----------------------------------------------------------------
            // Bulk data import – replace this section with your actual import logic.
            // Example: filling a range with new values.
            // -----------------------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            // Assume the data range starts at A2 and has 10 rows, 3 columns.
            for (int row = 0; row < 10; row++)
            {
                dataSheet.Cells[row + 1, 0].PutValue($"Item {row + 1}");
                dataSheet.Cells[row + 1, 1].PutValue(row * 10 + 5);   // Example numeric value
                dataSheet.Cells[row + 1, 2].PutValue(row * 8 + 3);    // Example numeric value
            }

            // After the data import, refresh all pivot‑based charts (including Waterfall
            // charts that use a pivot source) so they pick up the new data.
            workbook.Worksheets.RefreshAll();

            // If there are non‑pivot Waterfall charts, you can force a full recalculation
            // of the workbook, which also updates chart data ranges.
            workbook.CalculateFormula();

            // Save the updated workbook.
            workbook.Save("UpdatedWorkbook.xlsx");
        }
    }
}
