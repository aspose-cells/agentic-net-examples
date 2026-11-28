// Title: Enable leader lines for all series in the first chart of an existing XLSX workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells enable leader lines on chart series in an existing XLSX file | set HasLeaderLines property for chart series using Aspose.Cells .NET | load workbook, modify first chart leader lines, and save with Aspose.Cells | Aspose.Cells example for turning on leader lines in a loaded Excel chart | how to activate leader lines for all series of the first chart in a workbook using Aspose.Cells
// Tags: Aspose.Cells chart series leader line activation | open existing XLSX workbook Aspose.Cells | first worksheet chart manipulation C# | modify chart series properties Aspose.Cells | export modified workbook to XLSX using Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace LeaderLineModifier
{
    // The program opens an existing XLSX file, retrieves the first worksheet's first chart, enables leader lines for each series by setting HasLeaderLines to true, and saves the updated workbook to a new XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths.
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Validate that the input file exists.
            if (!File.Exists(inputPath))
            {
                Console.Error.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the existing workbook.
                Workbook workbook = new Workbook(inputPath);

                // Ensure the workbook contains at least one worksheet.
                if (workbook.Worksheets.Count == 0)
                {
                    Console.Error.WriteLine("Error: The workbook does not contain any worksheets.");
                    return;
                }

                // Use the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart.
                if (worksheet.Charts.Count == 0)
                {
                    Console.Error.WriteLine("Error: The first worksheet does not contain any charts.");
                    return;
                }

                // Retrieve the first chart.
                Chart chart = worksheet.Charts[0];
                Console.WriteLine($"Chart type: {chart.Type}");

                // Modify leader line settings for each series in the chart.
                foreach (Series series in chart.NSeries)
                {
                    try
                    {
                        // Enable leader lines.
                        series.HasLeaderLines = true;

                        // Note: Aspose.Cells does not expose direct color/width settings for leader lines.
                        // If needed, additional formatting can be applied via the series' shape properties.
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Warning: Unable to modify leader lines for a series. {ex.Message}");
                    }
                }

                // Save the modified workbook.
                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Error saving workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // Output any unexpected errors.
                Console.Error.WriteLine("An exception occurred while processing the workbook:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
