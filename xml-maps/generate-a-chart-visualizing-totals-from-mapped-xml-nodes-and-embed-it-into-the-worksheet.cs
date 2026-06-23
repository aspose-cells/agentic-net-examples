using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartFromXml
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // Assume XML data has been mapped to the worksheet.
            // For demonstration, we manually populate cells that would hold
            // the totals extracted from the XML nodes.
            // -----------------------------------------------------------------
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Total");

            // Sample data representing totals from mapped XML nodes
            sheet.Cells["A2"].PutValue("Item1");
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["A3"].PutValue("Item2");
            sheet.Cells["B3"].PutValue(230);
            sheet.Cells["A4"].PutValue("Item3");
            sheet.Cells["B4"].PutValue(95);
            sheet.Cells["A5"].PutValue("Item4");
            sheet.Cells["B5"].PutValue(180);

            // Add a column chart to the worksheet (using ChartCollection.Add)
            // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 1, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the chart.
            // The range includes both category (Item) and values (Total).
            // The second argument 'true' indicates that data is plotted by column.
            chart.SetChartDataRange("A1:B5", true);

            // Optional: set chart title and style
            chart.Title.Text = "Totals from XML Mapping";
            chart.Style = 2; // Built‑in style index

            // Calculate the chart layout before saving (optional but recommended)
            chart.Calculate();

            // Save the workbook with the embedded chart (lifecycle: save)
            workbook.Save("TotalsChartFromXml.xlsx", SaveFormat.Xlsx);
        }
    }
}