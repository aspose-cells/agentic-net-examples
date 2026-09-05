// Title: Check for primary and secondary category axes in a column chart using Chart.HasAxis with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, adds a column chart, and uses Chart.HasAxis to determine if the primary category axis exists. | Show how to call Chart.HasAxis with AxisType.Category for both primary and secondary axes in Aspose.Cells and output the results to the console. | Generate a console application that prints true/false for the presence of primary and secondary category axes in an Excel chart using Aspose.Cells .NET.
// Common Searches: aspnet cells chart.HasAxis category axis check c# example | how to detect if an Excel chart has a secondary category axis using Aspose.Cells | C# Aspose.Cells determine presence of primary category axis in column chart
// Tags: Chart.HasAxis category axis detection | primary category axis verification Aspose.Cells | secondary category axis check Aspose.Cells | column chart axis existence C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, populates sample data, adds a column chart, and uses Chart.HasAxis to check for both primary and secondary category axes, prints the results, and saves the file as ChartHasCategoryAxis.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.SetChartDataRange("A1:B4", true);

        // Use HasAxis to check for the presence of a category axis
        bool hasPrimaryCategoryAxis = chart.HasAxis(AxisType.Category, true);
        bool hasSecondaryCategoryAxis = chart.HasAxis(AxisType.Category, false);

        // Output the results
        Console.WriteLine("Primary Category Axis exists: " + hasPrimaryCategoryAxis);
        Console.WriteLine("Secondary Category Axis exists: " + hasSecondaryCategoryAxis);

        // Save the workbook
        workbook.Save("ChartHasCategoryAxis.xlsx");
    }
}
