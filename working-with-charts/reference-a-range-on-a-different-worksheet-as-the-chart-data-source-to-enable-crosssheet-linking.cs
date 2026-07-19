// Title: Aspose.Cells for .NET – Set a chart data range on a different worksheet
// Description: Creates a workbook with a "Data" sheet containing categories and values, adds a separate "ChartSheet", inserts a column chart, and uses SetChartDataRange("Data!A1:B5") to pull data from the first sheet. The chart title is linked to a cell on the data sheet, and the workbook is saved as CrossSheetChart.xlsx.
// Keywords: Aspose.Cells chart cross sheet | SetChartDataRange external worksheet | C# chart data source another sheet | link chart title to cell Aspose.Cells | column chart separate sheet .NET
// Common Searches: Aspose.Cells set chart data range on another worksheet | C# chart source from different sheet Aspose.Cells | link chart title to cell in another sheet Aspose.Cells | cross‑sheet chart example Aspose.Cells .NET
// Developer Intent: Create a chart on one worksheet that reads its data from a range located on a different worksheet and optionally bind the title to a source cell.
// Use Cases: Generate a dedicated chart sheet that visualizes sales figures stored on a master data sheet. | Build a reporting workbook where each chart resides on its own tab while sharing a common data source. | Automatically update chart titles by linking them to header cells on the data worksheet.
// AI Prompts: Show C# code to set a chart's data range to a range on another worksheet using Aspose.Cells. | Provide an example of linking a chart title to a cell on a different sheet so the title updates dynamically. | Explain how to create multiple charts on separate sheets that all reference the same data range in a source worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCrossSheetChart
{
    // Creates a workbook with a "Data" sheet containing categories and values, adds a separate "ChartSheet", inserts a column chart, and uses SetChartDataRange("Data!A1:B5") to pull data from the first sheet. The chart title is linked to a cell on the data sheet, and the workbook is saved as CrossSheetChart.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Sheet 1 – holds the data that will be used by the chart
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data (A1:B5)
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["A4"].PutValue("C");
            dataSheet.Cells["B4"].PutValue(30);
            dataSheet.Cells["A5"].PutValue("D");
            dataSheet.Cells["B5"].PutValue(40);

            // -------------------------------------------------
            // Sheet 2 – will contain the chart
            // -------------------------------------------------
            Worksheet chartSheet = workbook.Worksheets.Add("ChartSheet");

            // Add a column chart to the second sheet
            int chartIndex = chartSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = chartSheet.Charts[chartIndex];

            // Set the chart data range to reference the range on the first sheet
            // The address includes the sheet name (Data!A1:B5)
            chart.SetChartDataRange("Data!A1:B5", true);

            // Optional: link the chart title to a cell on the data sheet
            chart.Title.Text = "Cross‑Sheet Chart";
            chart.Title.LinkedSource = "='Data'!$A$1";

            // Save the workbook
            workbook.Save("CrossSheetChart.xlsx");
        }
    }
}
