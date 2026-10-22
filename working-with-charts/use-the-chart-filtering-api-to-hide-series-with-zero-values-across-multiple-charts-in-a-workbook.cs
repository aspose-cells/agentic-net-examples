// Title: Use Aspose.Cells for .NET to hide or delete chart series that contain only zero values across all worksheets in an Excel workbook
// AI Prompts: Generate C# code that iterates through every worksheet and chart, checks each series' data range, and removes the series if all cells are zero using Aspose.Cells. | Show how to create a method that evaluates a chart series' value range for numeric and string cells and filters out zero‑only series from multiple charts in a .NET application. | Provide a snippet that uses the Aspose.Cells chart filtering API to clean up Excel charts by deleting empty or zero‑valued series before saving the workbook.
// Common Searches: aspnet remove zero-valued series from Excel charts using Aspose.Cells | c# iterate all charts in workbook and delete series with all zeros | how to hide chart series with only zero data in multiple worksheets Aspose.Cells | filter out empty series from Excel chart programmatically .NET | Aspose.Cells chart series cleanup based on data values
// Tags: Aspose.Cells chart series zero-value removal | filter zero-valued series .NET Excel | iterate worksheets chart cleanup Aspose.Cells | chart series value range evaluation Aspose.Cells | remove empty series from Excel charts C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an Excel workbook, walks through each worksheet and its charts, examines the data range of every series, determines if all numeric or parsable string cells are zero, removes such series from the chart, and saves the cleaned workbook.
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

            // Load the workbook that contains the charts
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Iterate through each chart on the worksheet
                foreach (Chart chart in worksheet.Charts)
                {
                    // Process series in reverse order so we can safely remove them
                    for (int s = chart.NSeries.Count - 1; s >= 0; s--)
                    {
                        Series series = chart.NSeries[s];

                        // Get the address of the range that provides the series values
                        string valueRange = series.Values;

                        // Skip series without a value range
                        if (string.IsNullOrEmpty(valueRange))
                            continue;

                        // Create a Range object for the series values (the chart resides on this worksheet)
                        Aspose.Cells.Range range = worksheet.Cells.CreateRange(valueRange);

                        // Determine whether all numeric cells in the range are zero
                        bool allZero = true;
                        foreach (Cell cell in range)
                        {
                            if (cell.Type == CellValueType.IsNumeric)
                            {
                                if (cell.DoubleValue != 0)
                                {
                                    allZero = false;
                                    break;
                                }
                            }
                            else if (cell.Type == CellValueType.IsString)
                            {
                                if (double.TryParse(cell.StringValue, out double d))
                                {
                                    if (d != 0)
                                    {
                                        allZero = false;
                                        break;
                                    }
                                }
                                else
                                {
                                    // Non‑numeric string means the series is not all zero
                                    allZero = false;
                                    break;
                                }
                            }
                            // Blank or other cell types are treated as zero
                        }

                        // Remove the series if every value is zero
                        if (allZero)
                        {
                            chart.NSeries.RemoveAt(s);
                        }
                    }
                }
            }

            // Save the workbook with the filtered charts
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
