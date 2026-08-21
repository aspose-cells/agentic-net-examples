// Title: Create a Cross‑Sheet Column Chart in Aspose.Cells for .NET
// Description: Shows how to generate a workbook with a data worksheet and a separate chart worksheet, fill the data sheet with categories and values, use SetChartDataRange("DataSheet!$A$1:$B$4") to source chart data from another sheet, and link the chart title to a cell on the data sheet before saving.
// Keywords: Aspose.Cells cross sheet chart | SetChartDataRange .NET | chart data from another worksheet | link chart title to cell Aspose.Cells | C# Aspose.Cells chart example | column chart external range | Aspose.Cells chart source range | Aspose.Cells workbook chart linking
// Common Searches: Aspose.Cells reference another sheet for chart data | SetChartDataRange with sheet name Aspose.Cells | How to link chart title to a cell in Aspose.Cells | Create chart on one sheet using data from another sheet C# | Cross‑worksheet chart source Aspose.Cells | Aspose.Cells column chart from separate data sheet
// Developer Intent: Create a chart on one worksheet that pulls its data and optional title from a different worksheet using Aspose.Cells for .NET.
// Use Cases: Generate a sales column chart on a summary sheet while keeping raw data on a hidden data sheet, with the title linked to a cell for automatic updates. | Build a dashboard workbook where each chart on a main sheet references modular data ranges from separate worksheets, simplifying maintenance and data refresh. | Automate report creation that places charts on a cover page and isolates source data on protected sheets for security and clarity.
// AI Prompts: Write C# code with Aspose.Cells to add a pie chart on Sheet1 that uses the range Sheet2!$A$2:$B$5 as its data source and links the chart title to Sheet2!$C$1. | Explain the SetChartDataRange method for cross‑worksheet references in Aspose.Cells and demonstrate how to change the source range at runtime. | Provide step‑by‑step instructions to create multiple charts on a single sheet, each pulling data from different worksheets, using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to generate a workbook with a data worksheet and a separate chart worksheet, fill the data sheet with categories and values, use SetChartDataRange("DataSheet!$A$1:$B$4") to source chart data from another sheet, and link the chart title to a cell on the data sheet before saving.
class CrossSheetChartDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // First worksheet will host the chart
        Worksheet chartSheet = wb.Worksheets[0];
        chartSheet.Name = "ChartSheet";

        // Second worksheet contains the data source
        Worksheet dataSheet = wb.Worksheets.Add("DataSheet");
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("Value");
        dataSheet.Cells["A2"].PutValue("A");
        dataSheet.Cells["B2"].PutValue(10);
        dataSheet.Cells["A3"].PutValue("B");
        dataSheet.Cells["B3"].PutValue(20);
        dataSheet.Cells["A4"].PutValue("C");
        dataSheet.Cells["B4"].PutValue(30);

        // Add a chart to the first worksheet
        int chartIndex = chartSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = chartSheet.Charts[chartIndex];

        // Reference a range on the second worksheet as the chart data source
        chart.SetChartDataRange("DataSheet!$A$1:$B$4", true);

        // Optionally link the chart title to a cell on the data sheet
        dataSheet.Cells["D1"].PutValue("Sales Chart");
        chart.Title.LinkedSource = $"='{dataSheet.Name}'!$D$1";

        // Save the workbook
        wb.Save("CrossSheetChart.xlsx");
    }
}
