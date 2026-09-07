// Title: Remove the third chart from an Excel worksheet using Aspose.Cells Charts.RemoveAt in C#
// AI Prompts: Write C# code that creates a workbook, adds three column charts, verifies that at least three charts exist, and removes the chart at index 2 with Aspose.Cells. | Show how to safely delete a specific chart from a worksheet by checking the chart count before calling Charts.RemoveAt in Aspose.Cells.
// Common Searches: Aspose.Cells C# remove chart at specific index | how to delete the third chart in an Excel file using Aspose.Cells | Charts.RemoveAt example for conditional chart removal | C# code to check chart count before removing a chart with Aspose.Cells | remove a column chart from worksheet programmatically Aspose.Cells
// Tags: Aspose.Cells Charts.RemoveAt usage | remove chart by index C# | delete specific Excel chart Aspose.Cells | conditional chart removal in workbook | C# Aspose.Cells chart management

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, populates sample data, adds three column charts, checks that a third chart exists, removes it using Charts.RemoveAt(2), and saves the file as RemovedThirdChart.xlsx.
class RemoveThirdChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the charts
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add three charts to the worksheet
        for (int i = 0; i < 3; i++)
        {
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5 + i * 10, 0, 20 + i * 10, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
        }

        // Remove the third chart (zero‑based index 2) if it exists
        if (sheet.Charts.Count > 2)
        {
            sheet.Charts.RemoveAt(2);
        }

        // Save the workbook
        workbook.Save("RemovedThirdChart.xlsx");
    }
}
