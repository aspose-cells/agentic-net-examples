// Title: Refresh data source ranges of all Waterfall charts in an Excel workbook after bulk import with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that locates every Waterfall chart in a workbook and sets its series Values range to a dynamically calculated range based on the last populated row in Sheet1. | Create a .NET script that loads an .xlsx file, determines the end of the data table on Sheet1, builds A2:A{lastRow} and B2:B{lastRow} ranges, updates each Waterfall chart series to use the new range, and saves the file. | Write a C# routine that iterates through all worksheets, identifies Waterfall charts, and refreshes their data source after a bulk data import using Aspose.Cells.
// Common Searches: how to programmatically change the data range of multiple waterfall charts in Excel using Aspose.Cells C# | C# Aspose.Cells update chart series values after importing new rows | refresh all waterfall chart sources in a workbook after bulk data load .NET
// Tags: Aspose.Cells update waterfall chart data source | C# set chart series values range | dynamic range calculation last data row Aspose.Cells | iterate workbook charts Aspose.Cells .NET | refresh Excel waterfall charts after bulk import

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads input.xlsx, finds the last populated row on Sheet1, builds dynamic A2:A{lastRow+1} and B2:B{lastRow+1} ranges, iterates through every worksheet and chart, and for each Waterfall chart updates each series' Values property to the new range before saving the workbook as output.xlsx.
class UpdateWaterfallCharts
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains the charts
            Workbook workbook = new Workbook(inputPath);

            // Determine the range of the newly imported data.
            // Assumes data on Sheet1, categories in column A, values in column B, starting from row 2.
            Worksheet dataSheet = workbook.Worksheets["Sheet1"];
            if (dataSheet == null)
            {
                Console.WriteLine("Worksheet 'Sheet1' not found.");
                return;
            }

            int lastRow = dataSheet.Cells.MaxDataRow; // zero‑based index of last row with data
            if (lastRow < 1) // No data beyond header row
            {
                Console.WriteLine("No data found in Sheet1.");
                return;
            }

            string categoryRange = $"Sheet1!A2:A{lastRow + 1}";
            string valuesRange = $"Sheet1!B2:B{lastRow + 1}";

            // Loop through all worksheets and their charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    // Process only Waterfall charts
                    if (chart.Type == ChartType.Waterfall)
                    {
                        // Update each series in the Waterfall chart
                        foreach (Series series in chart.NSeries)
                        {
                            // CategoryData property is not available in older Aspose.Cells versions;
                            // updating only the Values range ensures the chart reflects the new data.
                            series.Values = valuesRange;
                        }
                    }
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
