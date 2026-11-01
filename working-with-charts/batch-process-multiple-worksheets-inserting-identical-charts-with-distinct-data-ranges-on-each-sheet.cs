// Title: Add identical column charts with worksheet‑specific titles to all sheets in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add the same column chart to every worksheet in a workbook | programmatically insert a chart on each sheet with a dynamic title using Aspose.Cells | batch create column charts for multiple worksheets in an existing Excel file C# | how to loop through worksheets and add charts with a fixed data range Aspose.Cells | set chart position and legend for each sheet using Aspose.Cells .NET API | add column chart to all worksheets and save workbook with Aspose.Cells
// Tags: Aspose.Cells add column chart to each worksheet | iterate worksheets create charts C# | fixed data range A1:B5 chart series Aspose.Cells | dynamic chart title using worksheet name | chart positioning rows columns Aspose.Cells | set legend position bottom Aspose.Cells chart | batch chart insertion Excel workbook .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook, loops through every worksheet, and inserts a column chart (rows 5‑20, columns 3‑15) using a fixed data range (A1:B5). Each chart receives a title that incorporates the worksheet name, a series name, and a bottom‑positioned legend. The modified workbook is then saved to a new file.
class BatchChartInsertion
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the data range for each worksheet (adjust as needed)
            string dataRange = "A1:B5";

            // Iterate through all worksheets and add a chart to each
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Add a column chart positioned from row 5, column 3 to row 20, column 15
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 3, 20, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart title
                chart.Title.Text = $"Sales Data - {sheet.Name}";
                chart.Title.IsVisible = true;

                // Add series using the defined data range; first column as categories, second as values
                int seriesIndex = chart.NSeries.Add(dataRange, true);
                chart.NSeries[seriesIndex].Name = $"Series_{sheet.Name}";

                // Optional appearance customizations
                chart.ShowLegend = true;
                chart.Legend.Position = LegendPositionType.Bottom;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
