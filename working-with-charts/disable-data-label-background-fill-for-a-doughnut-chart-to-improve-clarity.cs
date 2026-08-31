// Title: Disable background fill of data labels in a doughnut chart using Aspose.Cells for C#
// AI Prompts: Generate C# code with Aspose.Cells that creates a doughnut chart, shows data labels, and sets their background mode to Transparent. | Write a .NET example that adds a doughnut chart to a worksheet and removes the fill color of the chart's data labels. | Show how to configure DataLabels.BackgroundMode = BackgroundMode.Transparent for a doughnut chart in Aspose.Cells.
// Common Searches: Aspose.Cells C# doughnut chart data label background transparent | remove data label fill color from doughnut chart using Aspose.Cells .NET | set chart data labels to no background in Aspose.Cells C# example | how to hide data label background in Excel doughnut chart with Aspose.Cells
// Tags: Aspose.Cells doughnut chart label fill control | C# doughnut chart label transparency | disable chart label fill Aspose.Cells | transparent chart labels .NET | Aspose.Cells label styling C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // This example creates a workbook, adds sample data, inserts a doughnut chart, enables data labels, sets their background mode to Transparent, and saves the file as DoughnutNoLabelBackground.xlsx.
class DisableDataLabelBackground
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the doughnut chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["A4"].PutValue("Banana");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(50);
        worksheet.Cells["B3"].PutValue(30);
        worksheet.Cells["B4"].PutValue(20);

        // Add a doughnut chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Doughnut, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Set values
        chart.NSeries.CategoryData = "A2:A4";      // Set categories

        // Enable data labels and disable their background fill
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowValue = true;               // Show the numeric values
        dataLabels.BackgroundMode = BackgroundMode.Transparent; // No background fill

        // Save the workbook
        workbook.Save("DoughnutNoLabelBackground.xlsx");
    }
}
