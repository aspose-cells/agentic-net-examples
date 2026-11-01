// Title: Load a named range and bind it to a Waterfall chart series using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# bind named range to waterfall chart series | Load data from a defined name into a Waterfall chart with Aspose.Cells | Create a Waterfall chart from an Excel named range using Aspose.Cells .NET | Set chart series name from named range text in Aspose.Cells C# | Add a Waterfall chart to a workbook and use named range data in C#
// Tags: bind named range to waterfall chart Aspose.Cells | waterfall chart series from defined name Aspose.Cells | chart.NSeries.Add using named range Aspose.Cells | set series name from named range Aspose.Cells | add waterfall chart to worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, retrieves the named range "WaterfallData", adds a Waterfall chart to the first worksheet, binds the chart series to the named range with chart.NSeries.Add, assigns the series name from the named range's text, and saves the workbook with the new chart.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook containing the named range
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range (e.g., "WaterfallData")
                Name namedRange = workbook.Worksheets.Names["WaterfallData"];
                if (namedRange == null)
                {
                    Console.WriteLine("Named range 'WaterfallData' not found.");
                    return;
                }

                // The RefersTo property returns the address of the range (e.g., "Sheet1!$B$2:$B$10")
                string dataRange = namedRange.RefersTo;

                // Add a new Waterfall chart to the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 5, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Bind the series to the named range.
                // The second parameter (true) indicates that the range contains category (X) values.
                chart.NSeries.Add(dataRange, true);

                // Set the series name using the defined name's text
                chart.NSeries[0].Name = namedRange.Text;

                // Save the workbook with the new chart
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
