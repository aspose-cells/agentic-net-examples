// Title: Load an XLSX workbook and retrieve the first chart object from the first worksheet using Aspose.Cells for .NET
// AI Prompts: Open a .xlsx file with Aspose.Cells, navigate to the first worksheet, and obtain the Chart instance at position zero. | Read an Excel workbook, verify that charts are present, and print the chart’s type and name in C#.
// Common Searches: Aspose.Cells C# get first chart from worksheet in existing Excel file | How to read chart type and name from a loaded XLSX using Aspose.Cells .NET | Check chart collection count before accessing charts with Aspose.Cells | Retrieve chart objects from a workbook loaded with Aspose.Cells for .NET
// Tags: load xlsx workbook Aspose.Cells C# | access worksheet chart collection Aspose.Cells | extract initial chart from worksheet Aspose.Cells | output chart metadata C# | ensure chart collection non-empty Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // Loads an XLSX file, checks the first worksheet for charts, and if a chart exists prints its type and name; otherwise reports that no charts were found.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the XLSX workbook that contains a chart.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0).
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the first chart object from the worksheet's Charts collection.
            // Ensure that at least one chart exists to avoid an IndexOutOfRangeException.
            if (sheet.Charts.Count > 0)
            {
                Chart firstChart = sheet.Charts[0];

                // Example: output chart type and name.
                Console.WriteLine("Chart Type: " + firstChart.Type);
                Console.WriteLine("Chart Name: " + firstChart.Name);
            }
            else
            {
                Console.WriteLine("No charts found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine("An error occurred while processing the workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
