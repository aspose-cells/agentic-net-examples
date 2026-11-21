// Title: C# example: Load a named range and bind it to a Waterfall chart series using Aspose.Cells
// Common Searches: Aspose.Cells C# how to bind a named range to a waterfall chart series | example of creating a waterfall chart from a named range with Aspose.Cells for .NET | using NSeries.Add to reference a named range in a waterfall chart Aspose.Cells | load workbook and add waterfall chart programmatically in C# Aspose.Cells | error handling for missing named range when generating waterfall chart with Aspose.Cells
// Tags: waterfall chart series data source named range Aspose.Cells | add waterfall chart to worksheet using Aspose.Cells C# | NSeries.Add named range syntax Aspose.Cells | retrieve named range from workbook Aspose.Cells | handle missing named range exception Aspose.Cells | save workbook with chart Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The code loads an existing workbook, retrieves the named range "WaterfallData", adds a Waterfall chart to the first worksheet, binds the chart series to the named range using the NSeries.Add method with curly‑brace syntax, optionally sets the series name, and saves the workbook, while handling file‑not‑found and missing‑range errors.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string namedRange = "WaterfallData";

            try
            {
                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

                // Load the workbook that contains the named range with the data
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range that holds the waterfall data
                // Named ranges are stored in the Names collection of the workbook
                Name dataRange = workbook.Worksheets.Names[namedRange];
                if (dataRange == null)
                    throw new InvalidOperationException($"Named range '{namedRange}' does not exist in the workbook.");

                // Use the first worksheet (or create a new one) where the chart will be placed
                Worksheet sheet = workbook.Worksheets[0];

                // Add a Waterfall chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 5, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Bind the chart series to the named range.
                // The curly braces tell Aspose.Cells to use the named range directly.
                chart.NSeries.Add($"{{{namedRange}}}", true);

                // Optional: set a name for the series
                chart.NSeries[0].Name = "Waterfall Series";

                // Save the workbook with the newly created chart
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (FileNotFoundException ex)
            {
                Console.Error.WriteLine($"File error: {ex.Message}");
            }
            catch (InvalidOperationException ex)
            {
                Console.Error.WriteLine($"Operation error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
