// Title: Save each chart in an Excel workbook as a PNG file named with its worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, iterates over every worksheet and its charts, and writes each chart to a PNG file whose name combines the worksheet name and chart sequence number. | Show how to call Aspose.Cells' Chart.ToImage method in a .NET console app to export charts to PNG images while applying a custom filename pattern based on the parent worksheet.
// Common Searches: Aspose.Cells export chart to PNG with worksheet name in filename C# | How to loop through worksheets and save each chart as a separate image using Aspose.Cells | C# batch export Excel charts to PNG files Aspose.Cells example | Save Excel chart images with sheet-specific filenames using Aspose.Cells .NET | Chart.ToImage custom file naming pattern Aspose.Cells
// Tags: export charts to PNG with Aspose.Cells | chart.ToImage filename customization | iterate worksheets save chart images C# | batch chart image extraction Aspose.Cells | worksheet-based chart file naming .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartImageExporterApp
{
    // The program loads an Excel workbook, walks through each worksheet and its charts, and saves every chart as a PNG file whose name includes the worksheet name and chart index.
    class ChartImageExporter
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts on the current worksheet
                    for (int i = 0; i < sheet.Charts.Count; i++)
                    {
                        Chart chart = sheet.Charts[i];

                        // Build a file name using the worksheet name and chart index
                        string fileName = $"{sheet.Name}_Chart{i + 1}.png";

                        // Save the chart as a PNG image (default format)
                        chart.ToImage(fileName);
                        Console.WriteLine($"Chart saved: {fileName}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
