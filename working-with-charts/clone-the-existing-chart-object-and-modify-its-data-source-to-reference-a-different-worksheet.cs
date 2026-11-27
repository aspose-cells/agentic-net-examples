// Title: How to clone a chart and bind it to a different worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# clone chart and set new data range on another worksheet | Copy an existing chart to a second sheet and change its NSeries source in .NET | Duplicate a column chart and bind it to Sheet2 data using Aspose.Cells | Reassign chart data source after cloning chart in Aspose.Cells for C# | Create cloned chart with same style but different worksheet data in Aspose.Cells
// Tags: chart cloning Aspose.Cells C# | binding chart to another worksheet Aspose.Cells | copy chart style and legend Aspose.Cells | modify NSeries range C# | Aspose.Cells change chart data source

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook with two worksheets, fills each with category/value data, adds a column chart on the first sheet referencing its data, then clones the chart onto the second sheet, copies title, style, and legend settings, clears the default series, and reassigns the NSeries and category data to the second sheet before saving the workbook as 'ClonedChartDemo.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add a second worksheet
            Workbook workbook = new Workbook();

            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            // Worksheets.Add() returns the index of the new sheet
            int sheet2Idx = workbook.Worksheets.Add();
            Worksheet sheet2 = workbook.Worksheets[sheet2Idx];
            sheet2.Name = "Sheet2";

            // Populate data in Sheet1
            sheet1.Cells["A1"].PutValue("Category");
            sheet1.Cells["B1"].PutValue("Value");
            sheet1.Cells["A2"].PutValue("A");
            sheet1.Cells["A3"].PutValue("B");
            sheet1.Cells["A4"].PutValue("C");
            sheet1.Cells["B2"].PutValue(10);
            sheet1.Cells["B3"].PutValue(20);
            sheet1.Cells["B4"].PutValue(30);

            // Populate different data in Sheet2
            sheet2.Cells["A1"].PutValue("Category");
            sheet2.Cells["B1"].PutValue("Value");
            sheet2.Cells["A2"].PutValue("A");
            sheet2.Cells["A3"].PutValue("B");
            sheet2.Cells["A4"].PutValue("C");
            sheet2.Cells["B2"].PutValue(100);
            sheet2.Cells["B3"].PutValue(200);
            sheet2.Cells["B4"].PutValue(300);

            // Add a chart to Sheet1 and set its data source to Sheet1
            int originalChartIdx = sheet1.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart originalChart = sheet1.Charts[originalChartIdx];
            originalChart.NSeries.Add("Sheet1!$B$2:$B$4", true);
            originalChart.NSeries.CategoryData = "Sheet1!$A$2:$A$4";
            originalChart.Title.Text = "Original Chart (Sheet1 Data)";

            // Clone the chart onto Sheet2 by creating a new chart with the same type and position
            int clonedChartIdx = sheet2.Charts.Add(originalChart.Type, 5, 0, 20, 5);
            Chart clonedChart = sheet2.Charts[clonedChartIdx];

            // Copy basic properties (title, style, etc.) from the original chart
            clonedChart.Title.Text = "Cloned Chart (Sheet2 Data)";
            clonedChart.Style = originalChart.Style;
            clonedChart.ShowLegend = originalChart.ShowLegend;

            // Set the cloned chart's data source to reference Sheet2
            clonedChart.NSeries.Clear(); // Remove any default series
            clonedChart.NSeries.Add("Sheet2!$B$2:$B$4", true);
            clonedChart.NSeries.CategoryData = "Sheet2!$A$2:$A$4";

            // Save the workbook
            string outputPath = "ClonedChartDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
