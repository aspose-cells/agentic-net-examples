// Title: Save an updated chart workbook as XLSX while preserving formatting with Aspose.Cells for .NET
// Common Searches: how to modify a chart series and title in Aspose.Cells C# then save workbook as XLSX preserving formatting | Aspose.Cells .NET example for updating chart data range and exporting to XLSX without losing styles | C# code to change first chart series values and title and keep formatting when saving workbook with Aspose.Cells | save workbook with edited chart to XLSX using SaveFormat.Xlsx in Aspose.Cells for .NET | Aspose.Cells C# load existing XLSX, edit chart, and write new file keeping chart formatting
// Tags: chart series update Aspose.Cells C# | chart title modification Aspose.Cells | preserve chart formatting on XLSX save Aspose.Cells | load workbook and edit chart Aspose.Cells | SaveFormat.Xlsx implementation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing XLSX workbook, checks for a chart, updates the first series' data range and the chart title, then saves the workbook as a new XLSX file using SaveFormat.Xlsx, ensuring all chart formatting and workbook styles are retained.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Proceed only if the worksheet contains at least one chart
            if (sheet.Charts.Count > 0)
            {
                Chart chart = sheet.Charts[0];

                // Modify the first series values if a series exists
                if (chart.NSeries.Count > 0)
                {
                    // Define new data range for the series (adjust as needed)
                    string newValuesRange = "Sheet1!$B$2:$B$5";
                    chart.NSeries[0].Values = newValuesRange;
                }

                // Update the chart title
                chart.Title.Text = "Updated Chart Title";
            }

            // Save the workbook preserving all formatting and chart settings
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
