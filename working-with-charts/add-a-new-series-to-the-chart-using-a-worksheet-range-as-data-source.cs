// Title: Add a new series to an Aspose.Cells column chart from a worksheet range (C#)
// Description: Creates an in‑memory workbook, fills columns A‑C with categories and two data series, inserts a column chart, binds the first series and its X‑axis labels using range formulas, then adds a second series via NSeries.Add, and saves the file.
// Keywords: Aspose.Cells C# add chart series | NSeries.Add worksheet range | column chart multiple series Aspose.Cells | set chart category labels Aspose.Cells | chart data source range Aspose.Cells | Aspose.Cells chart example C#
// Common Searches: How to add another series to an existing Aspose.Cells chart in .NET | Aspose.Cells NSeries.Add using worksheet range | Set X‑axis category labels for a chart with Aspose.Cells C# | Add multiple data series to a column chart programmatically Aspose.Cells | Bind chart series to cell ranges Aspose.Cells
// Developer Intent: Add an additional data series to an existing Aspose.Cells chart by referencing a worksheet range as the source.
// Use Cases: Generate a column chart that shows two separate series from columns B and C while using column A for category labels. | Append extra series to a chart at runtime when new data columns become available. | Define X‑axis categories from a range and bind each series to its own column range within a single chart.
// AI Prompts: Write C# code to add a third series from column D to the chart using NSeries.Add with a worksheet range. | Show how to bind chart series to named ranges instead of explicit cell addresses in Aspose.Cells. | Explain how to programmatically update values of an existing chart series in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsAddSeriesExample
{
    // Creates an in‑memory workbook, fills columns A‑C with categories and two data series, inserts a column chart, binds the first series and its X‑axis labels using range formulas, then adds a second series via NSeries.Add, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in‑memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            // Column A – Category labels
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["A5"].PutValue("D");

            // Column B – First data series
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            // Column C – Second data series (will be added later)
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);
            sheet.Cells["C5"].PutValue(45);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the first series using a worksheet range as the data source
            // The range includes the header (Series1) and the values (B2:B5)
            chart.NSeries.Add("=Sheet1!$B$1:$B$5", true);

            // Define the category (X‑axis) labels range
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$5";

            // Add a second series to the same chart using another range
            // This demonstrates adding a new series after the chart has been created
            chart.NSeries.Add("=Sheet1!$C$1:$C$5", true);

            // Save the workbook to a file
            workbook.Save("AddSeriesUsingWorksheetRange.xlsx");
        }
    }
}
