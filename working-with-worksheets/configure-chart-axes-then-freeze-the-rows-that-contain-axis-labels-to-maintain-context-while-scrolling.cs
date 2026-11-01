// Title: Set chart axis titles and a fixed major unit, then freeze the header row in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Create a column chart, assign custom titles to the X and Y axes, set a fixed major unit for the value axis, and apply FreezePanes to keep the header row visible while scrolling with Aspose.Cells in C#. | Load an existing workbook, define the data range for a column chart, configure axis properties (titles, font size, major unit), freeze the first worksheet row, and save the updated file using Aspose.Cells for .NET.
// Common Searches: how to set a fixed major unit for a chart value axis using Aspose.Cells C# | freeze top row with axis labels in Excel workbook via Aspose.Cells .NET | configure chart axis titles programmatically in Aspose.Cells | add column chart and freeze panes in Aspose.Cells workbook C# example | Aspose.Cells set chart axis font size and major unit
// Tags: chart value axis major unit Aspose.Cells C# | freeze panes header row Aspose.Cells | column chart axis titles Aspose.Cells | set chart axis font size Aspose.Cells | load workbook add chart Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing workbook, adds a column chart with data from A2:A5 and B2:B5, sets the chart title, customizes the X and Y axis titles and font sizes, disables automatic major units and sets a fixed major unit of 5000 for the value axis, freezes the first worksheet row to keep axis labels visible, and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load an existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart to the worksheet (positioned from row 5, column 0 to row 20, column 10)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Sales Data";

            // Define the data range for the series (values) and the category (X) axis labels
            chart.NSeries.Add("B2:B5", true);          // Values from B2:B5
            chart.NSeries.CategoryData = "A2:A5";     // Labels from A2:A5

            // ----- Configure Value (Y) Axis -----
            chart.ValueAxis.Title.Text = "Revenue";
            chart.ValueAxis.Title.Font.Size = 12;
            // Example: set a fixed major unit
            chart.ValueAxis.IsAutomaticMajorUnit = false;
            chart.ValueAxis.MajorUnit = 5000;

            // ----- Configure Category (X) Axis -----
            chart.CategoryAxis.Title.Text = "Month";
            chart.CategoryAxis.Title.Font.Size = 12;
            // Note: IsBetweenCategories property is not available in this version; default behavior is used.

            // Freeze the top row so it stays visible while scrolling
            // FreezePanes(row, column, totalRows, totalColumns)
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
