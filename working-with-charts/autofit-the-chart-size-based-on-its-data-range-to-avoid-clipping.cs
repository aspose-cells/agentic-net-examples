// Title: Auto‑Fit Chart Size in Aspose.Cells for C# – Prevent Data Clipping
// Description: Demonstrates how to calculate the exact width and height an Excel chart needs, using Chart.Calculate and Chart.GetActualSize, then apply those dimensions to ChartObject so the chart displays all data without clipping.
// Keywords: Aspose.Cells chart auto size C# | Chart.GetActualSize example | Chart.Calculate Aspose.Cells | prevent chart clipping | dynamic chart resizing .NET | Excel chart size adjustment
// Common Searches: Aspose.Cells auto resize chart | C# get actual chart size Aspose | how to prevent chart clipping in Excel using Aspose | adjust chart width height programmatically Aspose.Cells | auto‑fit column chart Aspose.Cells C#
// Developer Intent: Programmatically set a chart’s width and height to the exact dimensions required for its data, eliminating clipping.
// Use Cases: Create column or line charts with large values and automatically size them for optimal display. | Generate Excel reports where chart dimensions must adapt to varying data ranges without manual tweaking. | Export workbooks that open in Excel with charts already sized correctly for readability.
// AI Prompts: Show C# code that uses Aspose.Cells Chart.Calculate and GetActualSize to auto‑size any chart. | Provide a snippet to set ChartObject.Width and Height based on the actual size returned by Aspose.Cells. | Explain the steps to ensure a chart updates its dimensions after data changes in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to calculate the exact width and height an Excel chart needs, using Chart.Calculate and Chart.GetActualSize, then apply those dimensions to ChartObject so the chart displays all data without clipping.
class AutoFitChartDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(1500); // large value to test auto‑fit

        // Add a column chart to the worksheet
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIdx];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Recalculate the chart layout so that size information is up‑to‑date
        chart.Calculate(); // uses Chart.Calculate rule

        // Get the size (width, height) that Excel would need to display the chart without clipping
        int[] actualSize = chart.GetActualSize(); // uses Chart.GetActualSize rule

        // Apply the calculated size to the chart's shape
        chart.ChartObject.Width = actualSize[0];
        chart.ChartObject.Height = actualSize[1];

        // Save the workbook
        workbook.Save("AutoFitChartDemo.xlsx"); // uses provided save rule
    }
}
