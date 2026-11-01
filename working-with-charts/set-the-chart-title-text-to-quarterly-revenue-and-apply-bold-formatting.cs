// Title: Set a bold chart title "Quarterly Revenue" in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing workbook or create a new one with Aspose.Cells, ensure a column chart exists on the first worksheet, assign the title text "Quarterly Revenue" to the chart, enable bold styling for the title font, and save the file. | Using C#, retrieve the first chart in the first worksheet of an Excel file with Aspose.Cells, change its title to "Quarterly Revenue", set the title font weight to bold, and write the updated workbook to disk.
// Common Searches: Aspose.Cells C# set chart title text to Quarterly Revenue | how to make chart title bold with Aspose.Cells for .NET | C# Aspose.Cells change column chart title and apply bold formatting | programmatically add a chart title and style it in an Excel file using Aspose.Cells | Aspose.Cells example for updating chart title font weight
// Tags: modify chart title Aspose.Cells | apply bold font to chart title Aspose.Cells | ensure column chart exists Aspose.Cells | access first worksheet chart Aspose.Cells | save workbook after chart changes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The code loads an existing workbook or creates a new one, guarantees a column chart is present on the first worksheet, sets the chart's title to "Quarterly Revenue", makes the title bold, and saves the updated workbook as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; otherwise create a minimal workbook with a chart.
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    // Create a new workbook with one worksheet.
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    ws.Name = "Sheet1";

                    // Add a simple chart to avoid index errors later.
                    int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 25, 10);
                    Chart placeholderChart = ws.Charts[chartIndex];
                    placeholderChart.Title.Text = "Placeholder";
                }

                // Get the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one chart; add one if none exist.
                if (worksheet.Charts.Count == 0)
                {
                    int chartIdx = worksheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
                    worksheet.Charts[chartIdx].Title.Text = "New Chart";
                }

                // Access the first chart.
                Chart chart = worksheet.Charts[0];

                // Set the chart title and apply bold formatting.
                chart.Title.Text = "Quarterly Revenue";
                chart.Title.Font.IsBold = true;

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
