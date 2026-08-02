// Title: Make doughnut chart data labels transparent in Aspose.Cells (C#)
// Description: Shows how to create an Excel workbook, add a doughnut chart, enable data labels, set their background to Transparent with Aspose.Cells for .NET, and save the file as DoughnutNoLabelBackground.xlsx.
// Keywords: Aspose.Cells | C# | doughnut chart | data labels | transparent background | BackgroundMode.Transparent | chart formatting | Excel label fill removal | chart visual clarity
// Common Searches: Aspose.Cells set doughnut chart label background transparent C# | remove data label fill Aspose.Cells chart | transparent data labels in Excel chart using Aspose.Cells | how to hide label background in doughnut chart Aspose.Cells | C# Aspose.Cells chart label formatting example
// Developer Intent: The developer wants to eliminate the colored fill behind data labels in a doughnut chart to produce a cleaner visual presentation.
// Use Cases: Generate a financial report where doughnut chart labels have no background color for a minimalist look. | Create a presentation slide with Excel charts that require transparent data labels to avoid visual clutter. | Automate dashboard exports where chart readability is improved by removing label fills.
// AI Prompts: Write C# code with Aspose.Cells that sets any chart's data label background to Transparent. | Explain the effect of BackgroundMode.Transparent on data label rendering in Aspose.Cells charts. | Provide a step‑by‑step guide to toggle the data label background fill on and off for a doughnut chart using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create an Excel workbook, add a doughnut chart, enable data labels, set their background to Transparent with Aspose.Cells for .NET, and save the file as DoughnutNoLabelBackground.xlsx.
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

        // Enable data labels for the series
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;

        // Disable the background fill of data labels for clarity
        series.DataLabels.BackgroundMode = BackgroundMode.Transparent;

        // Save the workbook
        workbook.Save("DoughnutNoLabelBackground.xlsx");
    }
}
