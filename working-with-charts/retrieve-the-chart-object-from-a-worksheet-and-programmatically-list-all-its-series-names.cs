// Title: List All Chart Series Names from a Worksheet with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, inserts a column chart, then accesses the chart object and iterates the NSeries collection to output each series' Name (or DisplayName if the Name is empty) before saving the file.
// Keywords: Aspose.Cells chart series names | C# retrieve chart series | Aspose.Cells NSeries enumeration | list chart series Aspose .NET | chart series Name fallback
// Common Searches: Aspose.Cells get chart series names C# | how to enumerate chart series in Aspose.Cells | retrieve chart object from worksheet Aspose | list series names Aspose.Cells chart | C# Aspose.Cells chart NSeries example
// Developer Intent: Obtain the chart object from a worksheet and output the name of each series.
// Use Cases: Validate that a workbook chart contains the expected series labels. | Populate a UI dropdown with chart series names for user selection. | Export series names to a reporting system or log file.
// AI Prompts: Generate C# code using Aspose.Cells that reads a worksheet chart and prints all series names, using DisplayName when Name is missing. | Show how to loop through the NSeries collection of an Aspose.Cells chart and handle empty Name properties. | Provide an example that lists chart series names and then saves the workbook to a given path.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart, then accesses the chart object and iterates the NSeries collection to output each series' Name (or DisplayName if the Name is empty) before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");

        worksheet.Cells["B1"].PutValue("Series 1");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        worksheet.Cells["C1"].PutValue("Series 2");
        worksheet.Cells["C2"].PutValue(15);
        worksheet.Cells["C3"].PutValue(25);
        worksheet.Cells["C4"].PutValue(35);

        // Add a chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data source for the chart series
        chart.NSeries.Add("B2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Iterate through all series in the chart and output their names
        for (int i = 0; i < chart.NSeries.Count; i++)
        {
            Series series = chart.NSeries[i];
            // Prefer the explicit Name property; fall back to DisplayName if Name is empty
            string seriesName = !string.IsNullOrEmpty(series.Name) ? series.Name : series.DisplayName;
            Console.WriteLine($"Series {i}: {seriesName}");
        }

        // Save the workbook
        workbook.Save("ChartSeriesNames.xlsx");
    }
}
