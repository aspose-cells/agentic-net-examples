// Title: Load an XLSX workbook from a file path and retrieve the first chart in the first worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to open a specified XLSX file, verify its existence, access the first worksheet, and return the first Chart object if any. | Write a C# console program with Aspose.Cells that loads a workbook, checks the chart collection of the first sheet, prints the chart type of the first chart, and includes proper exception handling for missing files.
// Common Searches: aspnet load xlsx file and get first chart Aspose.Cells | C# Aspose.Cells how to read chart type from first worksheet | example code to check chart collection in Excel workbook using Aspose.Cells .NET | retrieve chart object from worksheet with Aspose.Cells C# tutorial
// Tags: load xlsx workbook with Aspose.Cells | first worksheet chart collection Aspose.Cells | get first chart object Aspose.Cells | verify file existence before loading Aspose.Cells | exception handling for Excel chart retrieval Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Loads an XLSX workbook from a given path, confirms the file exists, accesses the first worksheet, checks for charts, and if present outputs the type of the first chart, with error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        // Path to the XLSX workbook
        string filePath = "input.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook from the file
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Check if the worksheet contains any charts
            if (worksheet.Charts.Count > 0)
            {
                // Obtain the first chart object
                Chart firstChart = worksheet.Charts[0];

                // Display the chart type
                Console.WriteLine("First chart type: " + firstChart.Type);
            }
            else
            {
                Console.WriteLine("No charts found in the first worksheet.");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
