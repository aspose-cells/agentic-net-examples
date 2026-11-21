// Title: Create and insert a column chart into the first worksheet of an existing Excel file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add column chart to existing workbook first sheet | How to programmatically insert a column chart in Excel with Aspose.Cells .NET | Aspose.Cells chart positioning rows 5 to 15 columns 0 to 5 example | Set data series range for column chart in Aspose.Cells C# | Save workbook after adding chart using Aspose.Cells
// Tags: column chart insertion Aspose.Cells .NET | first worksheet chart placement Excel | chart bounds rows 5-15 columns 0-5 Aspose.Cells | chart data series range A1:B5 Aspose.Cells | save workbook after chart addition Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example verifies the input file, loads it into a Workbook, accesses the first worksheet, adds a column chart positioned between rows 5‑15 and columns 0‑5, optionally defines the data series range (e.g., A1:B5), and saves the modified workbook as a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a column chart positioned from row 5, column 0 to row 15, column 5
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // (Optional) Define the data range for the chart series, e.g., cells A1:B5
                // chart.NSeries.Add("A1:B5", true);

                // Save the workbook with the new chart
                workbook.Save(outputPath);
                Console.WriteLine($"Chart added successfully. Workbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
