// Title: How to apply a consistent color palette to every chart series in an Excel workbook and save it as XLS using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells change chart series colors across all worksheets and export to XLS | apply same fill and border colors to every chart in a workbook with Aspose.Cells .NET | iterate through worksheets and charts to set a uniform palette in Aspose.Cells C# example | save modified workbook as Excel 97-2003 format using Aspose.Cells after updating chart colors | Aspose.Cells chart series color customization across multiple sheets C#
// Tags: apply uniform chart series colors Aspose.Cells | iterate worksheets charts Aspose.Cells C# | set series fill border color Aspose.Cells | save workbook as Excel97 XLS Aspose.Cells | chart color palette automation .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing XLSX workbook, defines a fixed array of colors, and iterates through each worksheet and each chart to assign those colors to the fill and border of every series. After processing all charts, the workbook is saved in the Excel 97‑2003 (XLS) format, with error handling for missing files and save failures.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define a uniform color palette to be applied to all chart series
            Color[] uniformPalette = new Color[]
            {
                Color.Blue,
                Color.Red,
                Color.Green,
                Color.Orange,
                Color.Purple
            };

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each chart on the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Apply colors from the uniform palette to each series in the chart
                        int colorIdx = 0;
                        foreach (Series series in chart.NSeries)
                        {
                            Color clr = uniformPalette[colorIdx % uniformPalette.Length];

                            // Set the fill color of the series
                            series.Area.ForegroundColor = clr;

                            // Set the border (line) color of the series
                            series.Border.Color = clr;

                            colorIdx++;
                        }
                    }
                    catch (Exception exChart)
                    {
                        Console.WriteLine($"Error processing chart '{chart.Name}': {exChart.Message}");
                    }
                }
            }

            // Save the modified workbook as an XLS file
            string outputPath = "output.xls";
            try
            {
                workbook.Save(outputPath, SaveFormat.Excel97To2003);
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
