// Title: Load an XLSX workbook and retrieve the first chart object using Aspose.Cells for .NET (C#)
// Common Searches: c# aspnet load xlsx workbook and read first chart with Aspose.Cells | how to access chart collection in Aspose.Cells and get chart name and type | retrieve first chart from worksheet using Aspose.Cells .NET example | aspocells validate chart count before accessing chart in C# | read chart properties from an Excel file with Aspose.Cells library
// Tags: aspocells workbook loading C# | aspocells worksheet chart collection access | aspocells initial chart object extraction | aspocells chart name and type reading | aspocells chart presence validation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample checks that the input.xlsx file exists, loads it into an Aspose.Cells Workbook, accesses the first worksheet, verifies that at least one chart is present, retrieves the first Chart object, and prints its Name and Type, with robust exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the XLSX workbook that contains a chart.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0).
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart.
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            // Retrieve the first chart object from the worksheet.
            Chart firstChart = worksheet.Charts[0];

            // Output basic information about the chart.
            Console.WriteLine("Chart Name: " + firstChart.Name);
            Console.WriteLine("Chart Type: " + firstChart.Type);
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
