// Title: Link Chart Series Data Labels to Source Cells' Number Format with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, populate category and series data, add a column chart, and link each series' data labels to a separate source column (D or E). The example uses the DataLabels.LinkedSource and NumberFormatLinked properties so the label text and number format follow the source cells, then saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | .NET | Excel chart data labels | LinkedSource property | NumberFormatLinked | custom label formatting | column chart | chart series formatting | Excel automation
// Common Searches: Aspose.Cells link data label to cell range | DataLabels.LinkedSource C# example | NumberFormatLinked Aspose.Cells chart | how to use custom data labels in Excel with Aspose | C# chart series label formatting from cells
// Developer Intent: Associate each chart series' data label number format with its dedicated source column to keep label appearance consistent with the underlying cells.
// Use Cases: Generate Excel reports where data labels show values with units or custom text stored in separate columns. | Maintain synchronized formatting between chart labels and source cells, so updates to cell styles automatically reflect on the chart. | Create multi‑series charts that require distinct label formats (e.g., different units or currency symbols) without manual label editing.
// AI Prompts: Convert the column chart in the example to a line chart while preserving data label linking. | Write code that determines the LinkedSource range dynamically based on the number of rows in the worksheet. | Explain the effect of NumberFormatLinked and show how to disable it for only the second series.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsSeriesDataLabelLinkDemo
{
    // Demonstrates how to create a workbook, populate category and series data, add a column chart, and link each series' data labels to a separate source column (D or E). The example uses the DataLabels.LinkedSource and NumberFormatLinked properties so the label text and number format follow the source cells, then saves the file as an Excel workbook.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column A: Categories
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            // Column B: Series 1 values
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Column C: Series 2 values
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(80);
            sheet.Cells["C3"].PutValue(110);
            sheet.Cells["C4"].PutValue(140);

            // Column D: Formatted labels for Series 1 (e.g., with units)
            sheet.Cells["D1"].PutValue("Series1 Labels");
            sheet.Cells["D2"].PutValue("120 units");
            sheet.Cells["D3"].PutValue("150 units");
            sheet.Cells["D4"].PutValue("180 units");

            // Column E: Formatted labels for Series 2
            sheet.Cells["E1"].PutValue("Series2 Labels");
            sheet.Cells["E2"].PutValue("80 units");
            sheet.Cells["E3"].PutValue("110 units");
            sheet.Cells["E4"].PutValue("140 units");

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Add first series (values from B2:B4) and link its data labels to D column
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries[0].XValues = "A2:A4";
            chart.NSeries[0].DataLabels.ShowValue = true;
            chart.NSeries[0].DataLabels.LinkedSource = "D2:D4";
            chart.NSeries[0].DataLabels.NumberFormatLinked = true; // link number format to source cells

            // Add second series (values from C2:C4) and link its data labels to E column
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries[1].XValues = "A2:A4";
            chart.NSeries[1].DataLabels.ShowValue = true;
            chart.NSeries[1].DataLabels.LinkedSource = "E2:E4";
            chart.NSeries[1].DataLabels.NumberFormatLinked = true; // link number format to source cells

            // Optional: customize appearance
            chart.Title.Text = "Series Data Labels Linked to Source Formats";
            chart.Legend.Position = LegendPositionType.Bottom;

            // Save the workbook
            workbook.Save("SeriesDataLabelsLinked.xlsx");
        }
    }
}
