// Title: Read Axis Labels After Chart.Calculate with Aspose.Cells (C#)
// Description: Creates a workbook, adds sample data, inserts a column chart, calls Chart.Calculate to generate axis ticks, then uses ValueAxis.GetAxisTexts() and CategoryAxis.GetAxisTexts() to retrieve the automatically computed value and category labels. The labels are printed to the console and the workbook is saved.
// Keywords: Aspose.Cells chart axis labels | Chart.Calculate GetAxisTexts C# | retrieve value axis texts Aspose.Cells | category axis tick labels after calculation | Aspose.Cells GetAxisTexts example
// Common Searches: how to get value axis labels from Aspose.Cells chart | Aspose.Cells GetAxisTexts after Chart.Calculate | C# read category axis tick marks Aspose.Cells | extract generated axis texts from a .NET chart
// Developer Intent: Obtain the automatically generated value and category axis labels of a chart after invoking Chart.Calculate.
// Use Cases: Verify axis labeling during automated testing. | Export chart tick labels to a report or logging system. | Drive dynamic formatting or annotation based on calculated labels.
// AI Prompts: Generate C# code that builds a line chart with Aspose.Cells, runs Chart.Calculate, and returns both value and category axis texts as string arrays. | Explain the relationship between Chart.Calculate and axis label generation in Aspose.Cells, and how GetAxisTexts extracts those labels. | Provide a step‑by‑step tutorial for extracting axis labels from any Aspose.Cells chart type after calculation using .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart, calls Chart.Calculate to generate axis ticks, then uses ValueAxis.GetAxisTexts() and CategoryAxis.GetAxisTexts() to retrieve the automatically computed value and category labels. The labels are printed to the console and the workbook is saved.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(8000);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(4000);
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B4"].PutValue(-8000);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the series and the category axis
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Calculate the chart to generate axis labels
        chart.Calculate();

        // Retrieve the generated labels for the value axis
        string[] valueAxisLabels = chart.ValueAxis.GetAxisTexts();

        // Retrieve the generated labels for the category axis
        string[] categoryAxisLabels = chart.CategoryAxis.GetAxisTexts();

        // Output the axis labels to the console
        Console.WriteLine("Value Axis Labels:");
        foreach (string label in valueAxisLabels)
        {
            Console.WriteLine(label);
        }

        Console.WriteLine("\nCategory Axis Labels:");
        foreach (string label in categoryAxisLabels)
        {
            Console.WriteLine(label);
        }

        // Save the workbook (optional)
        workbook.Save("AxisLabelsDemo.xlsx");
    }
}
