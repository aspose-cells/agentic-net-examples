// Title: Set Custom Colors for Stacked Bar Chart Series in Aspose.Cells for .NET (C#)
// Description: A complete C# example that creates a workbook, adds quarterly data, builds a stacked bar chart, and assigns a unique foreground color to each series using the Area.ForegroundColor property (Series.Style is not a direct color setter). The workbook is saved as an Excel file with the customized chart.
// Keywords: Aspose.Cells | C# | stacked bar chart | custom series colors | Series.Style | Area.ForegroundColor | chart series color .NET | Excel chart customization | Aspose.Cells example | set chart series color
// Common Searches: Aspose.Cells set series color stacked bar | C# change chart series color Aspose.Cells | How to use Series.Style for chart colors in Aspose.Cells | Assign custom colors to stacked bar chart series .NET | Aspose.Cells chart color customization example
// Developer Intent: Apply distinct colors to each series of a stacked bar chart using Aspose.Cells for .NET.
// Use Cases: Generate quarterly performance reports with brand‑specific colors for each series. | Automate Excel dashboard creation where series colors follow a corporate palette. | Create reusable chart templates that enforce consistent series coloring across multiple workbooks. | Produce client‑ready Excel files with visually differentiated stacked bar series.
// AI Prompts: Provide a C# example that sets custom colors for each series in a stacked bar chart using Aspose.Cells, demonstrating both Series.Style and Area.ForegroundColor approaches. | Explain why Series.Style cannot directly set series color in Aspose.Cells and how Area.ForegroundColor works as a workaround. | Show how to apply a predefined color palette to all series of a stacked bar chart before saving the workbook with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// A complete C# example that creates a workbook, adds quarterly data, builds a stacked bar chart, and assigns a unique foreground color to each series using the Area.ForegroundColor property (Series.Style is not a direct color setter). The workbook is saved as an Excel file with the customized chart.
class StackedBarCustomColors
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for a stacked bar chart
        // Categories
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");
        sheet.Cells["A5"].PutValue("Q4");

        // Series 1
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["B5"].PutValue(40);

        // Series 2
        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);
        sheet.Cells["C5"].PutValue(45);

        // Series 3
        sheet.Cells["D1"].PutValue("Series3");
        sheet.Cells["D2"].PutValue(12);
        sheet.Cells["D3"].PutValue(22);
        sheet.Cells["D4"].PutValue(32);
        sheet.Cells["D5"].PutValue(42);

        // Add a stacked bar chart
        int chartIdx = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 20, 15);
        Chart chart = sheet.Charts[chartIdx];

        // Set the data range for the series (B2:D5) and categories (A2:A5)
        chart.NSeries.Add("B2:D5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Assign custom colors to each series using the Area.ForegroundColor property
        // (Series.Style is not a direct property; the visual color is controlled via the Area)
        chart.NSeries[0].Area.ForegroundColor = Color.FromArgb(79, 129, 189); // Series1 color
        chart.NSeries[1].Area.ForegroundColor = Color.FromArgb(192, 80, 77); // Series2 color
        chart.NSeries[2].Area.ForegroundColor = Color.FromArgb(155, 187, 89); // Series3 color

        // Save the workbook
        workbook.Save("StackedBarCustomColors.xlsx");
    }
}
