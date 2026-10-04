// Title: Synchronize Excel chart series data label number format with its source column using Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates over all charts in a workbook, reads the number format of the first column in each series' source range, and sets series.DataLabels.NumberFormat to that format with Aspose.Cells. | Show how to enable data labels for every series in a worksheet chart and apply the source cells' custom or built‑in number format to the labels using Aspose.Cells for .NET. | Provide a snippet that logs any errors while processing chart series and ensures the workbook is saved after updating data label formats with Aspose.Cells.
// Common Searches: Aspose.Cells how to copy number format from source cells to chart data labels in C# | C# set Excel chart series data label format based on source column using Aspose | Apply cell number format to Excel chart data labels programmatically with Aspose.Cells | Synchronize chart data label formatting with underlying worksheet data in .NET
// Tags: chart series data label formatting Aspose.Cells | extract source column number format C# | map cell number format to chart labels | enable data labels for worksheet charts Aspose | consistent chart formatting with source cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program loads an Excel workbook, iterates through each chart and its series, enables data labels, extracts the number format from the first column of the series' source range, assigns that format to the series' DataLabels.NumberFormat, and saves the modified workbook.
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
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the chart is on the first worksheet; adjust index if needed
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all charts on the worksheet
            foreach (Chart chart in worksheet.Charts)
            {
                // Process each series in the chart
                foreach (Series series in chart.NSeries)
                {
                    try
                    {
                        // Enable data labels for the series
                        series.DataLabels.ShowValue = true;

                        // The series.Values property contains the source range formula (e.g., "Sheet1!$B$2:$B$5")
                        string sourceRangeFormula = series.Values;
                        if (string.IsNullOrEmpty(sourceRangeFormula))
                            continue;

                        // Extract the address part (remove sheet name if present)
                        string address = sourceRangeFormula;
                        int exclPos = sourceRangeFormula.IndexOf('!');
                        if (exclPos >= 0 && exclPos < sourceRangeFormula.Length - 1)
                            address = sourceRangeFormula.Substring(exclPos + 1);

                        // Create a Range object from the address (fully qualified to avoid ambiguity)
                        Aspose.Cells.Range sourceRange = worksheet.Cells.CreateRange(address);

                        // Determine the first column of the source range
                        int sourceColumnIndex = sourceRange.FirstColumn;

                        // Retrieve the style of the first cell in the source column
                        Style sourceStyle = worksheet.Cells[sourceRange.FirstRow, sourceColumnIndex].GetStyle();

                        // Extract the number format string from the source style
                        string numberFormat = !string.IsNullOrEmpty(sourceStyle.Custom)
                            ? sourceStyle.Custom
                            : sourceStyle.Number.ToString();

                        // Apply the extracted number format to the data labels of the series
                        series.DataLabels.NumberFormat = numberFormat;
                    }
                    catch (Exception exSeries)
                    {
                        Console.WriteLine($"Error processing series: {exSeries.Message}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
