// Title: Modify every chart’s title and first series range in an Excel workbook and save as XLSX while preserving chart controls using Aspose.Cells for .NET
// AI Prompts: Write a C# program that opens an existing .xlsx file with Aspose.Cells, loops through all worksheets and charts, sets each chart’s Title.Text to a custom string, updates the first series Values range, and then saves the workbook as XLSX keeping all chart objects and controls intact. | Generate a .NET code snippet that demonstrates how to retain chart formatting, controls, and embedded objects when saving a modified workbook after changing chart properties with Aspose.Cells.
// Common Searches: how to keep Excel chart objects when saving with Aspose.Cells C# | update chart title and series range for all sheets using Aspose.Cells .NET | preserve chart controls after modifying charts in a workbook with Aspose.Cells | Aspose.Cells iterate through charts and change data source in C# | save workbook with modified charts without losing formatting Aspose.Cells
// Tags: modify chart title Aspose.Cells | update chart series values range C# | save workbook as XLSX preserving charts Aspose.Cells | iterate worksheets and charts Aspose.Cells .NET | retain chart controls after save Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing XLSX workbook, iterates through each worksheet and chart, changes the chart title and the first series data range, then saves the workbook as XLSX while preserving all chart objects, controls, and formatting using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at path '{inputPath}'.");
            return;
        }

        try
        {
            // Load the existing workbook that contains charts
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts on the current worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Update the chart title
                    chart.Title.Text = "Updated Chart Title";

                    // Change the data range of the first series, if it exists
                    if (chart.NSeries.Count > 0)
                    {
                        var firstSeries = chart.NSeries[0];
                        // Set a new values range (adjust the range as needed)
                        firstSeries.Values = "=Sheet1!$B$2:$B$5";
                    }

                    // Aspose.Cells automatically retains all chart objects, controls, and formatting
                }
            }

            // Save the workbook, preserving all chart controls and modifications
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
