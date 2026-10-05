// Title: Binding a column chart to a merged‑cell range with Aspose.Cells SetChartDataRange in C#
// AI Prompts: Create a C# workbook with Aspose.Cells, merge the value cells, and bind a column chart to that merged range using SetChartDataRange. | Demonstrate adding a series to an Aspose.Cells chart when the numeric values are stored in a merged block of cells. | Produce an Excel file where the chart title and series correctly reflect data from a merged range using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set chart data range that includes merged cells | How to display merged cell values in an Excel chart using Aspose.Cells | Vertical orientation chart binding with merged data using Aspose.Cells SetChartDataRange example | Column chart series from merged cells Aspose.Cells C# tutorial
// Tags: Aspose.Cells SetChartDataRange merged area | C# column chart data binding Aspose.Cells | Excel chart series from merged cell region .NET | Aspose.Cells chart merged data handling | SetChartDataRange vertical orientation example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, adds category and value data, merges cells B3:B4, inserts a column chart, binds the chart to the range A1:B4 with vertical orientation using SetChartDataRange, defines the series from B2:B4, sets a chart title, and saves the file as ChartWithMergedCells.xlsx.
class SetChartDataRangeWithMergedCells
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data in column A (categories)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            // Populate data in column B (values)
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(120);
            // Merge B3 and B4 to simulate merged data cell
            sheet.Cells.Merge(2, 1, 2, 1); // Merges B3:B4 (rows 2-3, column 1)
            sheet.Cells["B3"].PutValue(150); // Value for the merged cell

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Bind the chart to the data range that includes merged cells.
            // The second argument 'true' indicates vertical orientation (categories in first column).
            chart.SetChartDataRange("A1:B4", true);

            // Define the series explicitly (category axis uses column A, values use column B)
            // The second parameter 'true' indicates that the first column contains category labels.
            chart.NSeries.Add("B2:B4", true);

            // Set the chart title
            chart.Title.Text = "Sales by Quarter";

            // Save the workbook
            string outputPath = "ChartWithMergedCells.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
