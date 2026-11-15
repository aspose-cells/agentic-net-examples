// Title: How to retrieve a chart from a worksheet and list all its series names using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get first chart from a worksheet and read its series names | How to list series names of an Excel chart with Aspose.Cells .NET | Example code to iterate over NSeries collection in Aspose.Cells | Retrieve chart object and print each series title using Aspose.Cells C#
// Tags: Aspose.Cells retrieve chart object | Aspose.Cells enumerate chart series names | C# read Excel chart series | Aspose.Cells NSeries iteration | Excel chart series extraction .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample loads an existing workbook (or creates a new one), accesses the first worksheet, checks for charts, retrieves the first chart, iterates through its NSeries collection, prints each series name to the console, and saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Load workbook if the file exists; otherwise create a new workbook.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Ensure at least one worksheet exists.
            if (workbook.Worksheets.Count == 0)
                workbook.Worksheets.Add("Sheet1");

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Check for charts in the worksheet.
            if (worksheet.Charts.Count > 0)
            {
                // Retrieve the first chart.
                Chart chart = worksheet.Charts[0];

                // Iterate through all series in the chart.
                for (int i = 0; i < chart.NSeries.Count; i++)
                {
                    string seriesName = chart.NSeries[i].Name;
                    Console.WriteLine($"Series {i + 1}: {seriesName}");
                }
            }
            else
            {
                Console.WriteLine("No charts found in the worksheet.");
            }

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook (optional if no changes were made).
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
