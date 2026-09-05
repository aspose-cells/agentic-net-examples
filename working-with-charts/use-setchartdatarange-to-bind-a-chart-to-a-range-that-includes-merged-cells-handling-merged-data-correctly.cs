// Title: Bind a column chart to a merged‑cell range using SetChartDataRange in Aspose.Cells for .NET
// AI Prompts: Create an Excel workbook in C# with a merged category column and add a column chart that reads the merged cells correctly using SetChartDataRange. | Write C# code to bind a column chart to a range that contains merged cells, ensuring the series are plotted vertically with Aspose.Cells. | Show how to handle merged header cells when setting the data source for a chart via Aspose.Cells' SetChartDataRange method.
// Common Searches: Aspose.Cells C# set chart data range with merged cells | bind chart to merged cell range using SetChartDataRange | column chart showing merged category column Aspose.Cells .NET | how to include merged cells in chart data source Aspose.Cells
// Tags: setchartdatarange merged range | aspocells column chart binding | c# excel chart merged handling | aspocells vertical series orientation | excel chart data source merged area

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExample
{
    // The example creates a new workbook, merges cells A2:A4 to represent a single category, fills values in column B, adds a column chart, and uses SetChartDataRange("A1:B4", true) to bind the chart to a range that includes the merged cells. The chart title is set and the workbook is saved as ChartWithMergedCells.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate header cells
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                // Merge cells A2:A4 to represent a single category spanning multiple rows
                sheet.Cells.Merge(1, 0, 3, 1); // Row 1, Column 0, 3 rows, 1 column (A2:A4)
                sheet.Cells["A2"].PutValue("Group1");

                // Fill values for each sub‑row
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 7);
                Chart chart = sheet.Charts[chartIndex];

                // Bind the chart to the range that includes the merged cells.
                // The second argument (true) indicates that the data series are plotted vertically.
                chart.SetChartDataRange("A1:B4", true);

                // Set the chart title
                chart.Title.Text = "Values by Category";

                // Save the workbook
                workbook.Save("ChartWithMergedCells.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
