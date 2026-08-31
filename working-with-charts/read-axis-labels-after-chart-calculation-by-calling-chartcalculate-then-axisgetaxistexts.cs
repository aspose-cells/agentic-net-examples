// Title: Use Aspose.Cells in C# to calculate a chart and read generated value and category axis labels
// AI Prompts: Generate C# code that builds a column chart with Aspose.Cells, calls Chart.Calculate(), and returns the array of strings from ValueAxis.GetAxisTexts(). | Write a C# snippet that extracts both value and category axis texts from a calculated Aspose.Cells chart using GetAxisTexts() and prints them. | Show how to save a workbook after retrieving calculated axis labels from a chart in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells how to get value axis labels after chart.Calculate in C# | C# Aspose.Cells retrieve category axis texts from a chart | example of using GetAxisTexts with Aspose.Cells chart objects | read calculated axis labels from column chart Aspose.Cells .NET | Chart.Calculate then ValueAxis.GetAxisTexts Aspose.Cells sample
// Tags: Chart.Calculate axis label extraction Aspose.Cells | ValueAxis.GetAxisTexts C# | CategoryAxis.GetAxisTexts Aspose.Cells | column chart axis labels .NET | retrieve chart axis texts Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart, calls Chart.Calculate() to generate axis labels, then uses ValueAxis.GetAxisTexts() and CategoryAxis.GetAxisTexts() to obtain the calculated labels, prints them, and saves the workbook.
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

        // Calculate the chart so that axis labels are generated
        chart.Calculate();

        // Retrieve the calculated labels for the value axis
        string[] valueAxisLabels = chart.ValueAxis.GetAxisTexts();

        // Retrieve the calculated labels for the category axis
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

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("AxisLabelsDemo.xlsx");
    }
}
