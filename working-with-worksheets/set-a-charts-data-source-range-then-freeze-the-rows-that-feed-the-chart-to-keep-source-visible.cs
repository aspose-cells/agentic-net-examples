// Title: Aspose.Cells for .NET – Set Chart Data Range and Freeze Source Rows (C#)
// Description: C# example that creates a workbook, fills rows 1‑5 with sample data, adds a column chart linked to range A1:B5, freezes the first five rows using FreezePanes, and saves the file as ChartWithFrozenSource.xlsx.
// Keywords: Aspose.Cells | C# chart data range | SetChartDataRange | FreezePanes | freeze worksheet rows | column chart Aspose.Cells | Excel chart source visibility | Aspose.Cells example GitHub | worksheet freeze panes API | Aspose.Cells for .NET tutorial
// Common Searches: Aspose.Cells set chart data range C# | How to freeze rows that feed a chart in Aspose.Cells | FreezePanes after creating a chart Aspose.Cells | C# example chart with frozen source rows | SetChartDataRange with headers Aspose.Cells
// Developer Intent: Create a column chart, bind it to a specific range, and lock the source rows so they stay visible while scrolling.
// Use Cases: Financial dashboards where the data table above a chart must remain in view. | Sales reports that combine a chart with a frozen data grid for easy comparison. | Automated Excel exports that protect chart source data from accidental scrolling.
// AI Prompts: Generate C# code to freeze both rows and columns that contain a chart’s source data using Aspose.Cells. | Show how to bind a chart to a range without headers in Aspose.Cells for .NET. | Explain the steps to apply FreezePanes after adding multiple charts so each chart’s source rows stay visible.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartFreezeDemo
{
    // C# example that creates a workbook, fills rows 1‑5 with sample data, adds a column chart linked to range A1:B5, freezes the first five rows using FreezePanes, and saves the file as ChartWithFrozenSource.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart (rows 1‑5)
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B4"].PutValue(30);
            worksheet.Cells["A5"].PutValue("D");
            worksheet.Cells["B5"].PutValue(40);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the chart's data source range (including headers)
            chart.SetChartDataRange("A1:B5", true);

            // Freeze the rows that contain the chart source data (rows 1‑5)
            // FreezePanes(string cellName, int freezedRows, int freezedColumns)
            // Cell "A6" is the first row below the data; freeze 5 rows above it.
            worksheet.FreezePanes("A6", 5, 0);

            // Save the workbook
            workbook.Save("ChartWithFrozenSource.xlsx");
        }
    }
}
