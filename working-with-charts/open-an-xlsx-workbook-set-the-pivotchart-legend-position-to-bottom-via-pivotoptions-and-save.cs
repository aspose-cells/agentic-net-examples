// Title: Set the legend position to Bottom for every chart in an existing XLSX workbook using Aspose.Cells for .NET
// AI Prompts: Open an XLSX workbook with Aspose.Cells, traverse each worksheet and chart, assign Legend.Position = LegendPositionType.Bottom, then write the file to a new location. | Create C# code that loads a spreadsheet, checks for a legend on every chart, moves the legend to the bottom edge, and saves the modified workbook.
// Common Searches: Aspose.Cells C# change legend placement of all charts in an Excel file | bulk update Excel chart legends to bottom using .NET library | programmatically set chart legend to lower position in existing XLSX | iterate over workbook charts and adjust legend location with Aspose.Cells | how to reposition chart legends in a saved Excel workbook via C#
// Tags: Aspose.Cells set chart legend bottom | C# modify Excel chart legend position | iterate workbook charts Aspose.Cells | update legend placement XLSX via .NET | save workbook after chart legend change

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing XLSX file with Aspose.Cells, iterates through each worksheet and its charts, sets the legend position to Bottom when a legend is present, ensures the output directory exists, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through worksheets and their charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Set legend position to Bottom for all charts
                        if (chart.Legend != null)
                        {
                            chart.Legend.Position = LegendPositionType.Bottom;
                        }
                    }
                    catch (Exception exChart)
                    {
                        Console.WriteLine($"Error processing chart on sheet '{sheet.Name}': {exChart.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception exSave)
            {
                Console.WriteLine($"Error saving workbook: {exSave.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
