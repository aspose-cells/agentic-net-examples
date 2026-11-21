// Title: Create a sparkline in cell L8 from the named range 'SalesData' using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# sparkline from a named range example | how to set a sparkline data source to a defined name in Aspose.Cells | place a sparkline at cell L8 programmatically with Aspose.Cells .NET | binding Excel named range to a sparkline using Aspose.Cells C# code | generate line sparkline from SalesData named range in a workbook
// Tags: create sparkline from named range Aspose.Cells | set sparkline location cell L8 C# | bind named range to sparkline Aspose.Cells .NET | Aspose.Cells sparkline data source Excel | add line sparkline programmatically C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an existing workbook, uses the named range 'SalesData' as the data source for a sparkline, positions the sparkline in cell L8, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that contains the named range 'SalesData'
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Add a line chart that uses the named range 'SalesData' as its source
            // Chart will be placed from row 5, column 0 to row 15, column 10 (adjust as needed)
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 15, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add the named range as a series; the second parameter indicates that the range contains values
            chart.NSeries.Add("SalesData", true);
            chart.Title.Text = "Sales Data Sparkline Equivalent";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
