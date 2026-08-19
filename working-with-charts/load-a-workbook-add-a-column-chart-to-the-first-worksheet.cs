// Title: Insert a Column Chart into the First Worksheet of an Existing Excel File using Aspose.Cells for .NET (C#)
// Description: Loads "input.xlsx", accesses the first worksheet, adds a column chart positioned from row 5‑column 1 to row 20‑column 8, sets the series to A1:B5 and categories to A1:A5, then saves the workbook as "output.xlsx". Demonstrates ChartCollection.Add, NSeries configuration, and Workbook.Save with Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | .NET | add column chart | Excel chart creation | load workbook | save workbook | ChartCollection.Add | NSeries range | first worksheet chart
// Common Searches: how to add a column chart to an existing Excel file using Aspose.Cells C# | Aspose.Cells chart data range example | C# create chart on first worksheet with Aspose.Cells | Aspose.Cells add chart and save workbook | set NSeries range for column chart Aspose.Cells
// Developer Intent: Add a column chart to the first sheet of a loaded workbook and persist the changes.
// Use Cases: Generate a sales column chart in a template report and output a ready‑to‑share workbook. | Automate monthly performance dashboards by loading data files, inserting a column chart, and saving each result. | Enhance an existing financial model with a new column chart that reflects refreshed data ranges.
// AI Prompts: Write C# code with Aspose.Cells that loads "input.xlsx", adds a clustered column chart to the first worksheet (rows 5‑20, columns 1‑8), sets series to "A1:B5" and categories to "A1:A5", then saves as "output.xlsx". | Explain how to modify the chart type, position, and data source for a chart created with Aspose.Cells in C#. | Provide a step‑by‑step tutorial for loading a workbook, inserting a column chart, configuring its NSeries, and saving the file using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExample
{
    // Loads "input.xlsx", accesses the first worksheet, adds a column chart positioned from row 5‑column 1 to row 20‑column 8, sets the series to A1:B5 and categories to A1:A5, then saves the workbook as "output.xlsx". Demonstrates ChartCollection.Add, NSeries configuration, and Workbook.Save with Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook from file (uses Workbook(string) constructor)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart to the worksheet (uses ChartCollection.Add method)
            // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 1, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // (Optional) Define the data range for the chart
            // Adjust the range according to the actual data in the worksheet
            chart.NSeries.Add("A1:B5", true); // Series values
            chart.NSeries.CategoryData = "A1:A5"; // Category labels

            // Save the modified workbook (uses Workbook.Save(string) method)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Chart added and workbook saved to '{outputPath}'.");
        }
    }
}
