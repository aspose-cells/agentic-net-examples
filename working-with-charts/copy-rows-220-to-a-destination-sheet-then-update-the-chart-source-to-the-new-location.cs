// Title: Aspose.Cells .NET: Copy Rows 2‑20 to a New Sheet and Reset Chart Source
// Description: Demonstrates how to use Aspose.Cells for .NET to fill rows 2‑20 on a "Source" worksheet, copy that block to a "Destination" worksheet with Cells.CopyRows, create a column chart on the new sheet, and reassign its data range using Chart.SetChartDataRange before saving the workbook.
// Keywords: Aspose.Cells copy rows | C# Excel chart update | SetChartDataRange example | Cells.CopyRows method | Aspose.Cells chart source range | .NET workbook manipulation | duplicate data sheet Aspose | Excel chart data range C# | Aspose.Cells tutorial
// Common Searches: Aspose.Cells copy specific rows to another worksheet C# | How to change chart data range after copying rows in Aspose.Cells | SetChartDataRange usage with copied data Aspose.Cells .NET | CopyRows example for Excel charts Aspose | Update chart source after moving data Aspose.Cells
// Developer Intent: Transfer rows 2‑20 from a source sheet to a destination sheet and point a chart to the newly copied range.
// Use Cases: Generate a summary sheet that mirrors a data segment and displays its own chart. | Create a printable report where the chart reflects only the copied rows. | Separate analysis view by moving a data block to another worksheet while preserving visual representation.
// AI Prompts: Show C# code that copies rows 2‑20 from one worksheet to another with Aspose.Cells and updates a chart to reference the new range. | Explain how to combine Cells.CopyRows and Chart.SetChartDataRange to duplicate data and adjust chart sources in Aspose.Cells for .NET. | Provide steps to reposition a chart after copying rows and resetting its data range using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCopyRowsAndUpdateChart
{
    // Demonstrates how to use Aspose.Cells for .NET to fill rows 2‑20 on a "Source" worksheet, copy that block to a "Destination" worksheet with Cells.CopyRows, create a column chart on the new sheet, and reassign its data range using Chart.SetChartDataRange before saving the workbook.
    class Program
    {
        static void Main()
        {
            // ---------- Create a workbook and source worksheet ----------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet srcSheet = workbook.Worksheets[0];           // source sheet (default name "Sheet1")
            srcSheet.Name = "Source";

            // Populate sample data in rows 2‑20 (Excel rows 2‑20 correspond to zero‑based indices 1‑19)
            // Column A: Category, Column B: Value
            srcSheet.Cells["A1"].PutValue("Category");
            srcSheet.Cells["B1"].PutValue("Value");
            for (int i = 1; i <= 19; i++)                         // i = 1..19 -> rows 2‑20
            {
                srcSheet.Cells[i, 0].PutValue($"Item {i}");
                srcSheet.Cells[i, 1].PutValue(i * 10);
            }

            // Add a chart on the source sheet that uses the source data range A2:B20
            int srcChartIdx = srcSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart srcChart = srcSheet.Charts[srcChartIdx];
            srcChart.NSeries.Add("=Source!$A$2:$B$20", true);      // set data source
            srcChart.Title.Text = "Source Chart";

            // ---------- Add destination worksheet ----------
            Worksheet destSheet = workbook.Worksheets.Add("Destination");

            // ---------- Copy rows 2‑20 from source to destination ----------
            // sourceRowIndex = 1 (row 2), destinationRowIndex = 1 (row 2), rowNumber = 19 (rows 2‑20)
            destSheet.Cells.CopyRows(srcSheet.Cells, 1, 1, 19);

            // ---------- Create a chart on the destination sheet ----------
            int destChartIdx = destSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart destChart = destSheet.Charts[destChartIdx];
            // Use SetChartDataRange to point to the copied data (A2:B20 on Destination sheet)
            destChart.SetChartDataRange("A2:B20", true);
            destChart.Title.Text = "Destination Chart";

            // Optional: move the destination chart to a different position
            destChart.Move(10, 2, 20, 8);   // topRow, leftColumn, bottomRow, rightColumn

            // ---------- Save the workbook ----------
            workbook.Save("RowsCopiedAndChartUpdated.xlsx");
        }
    }
}
