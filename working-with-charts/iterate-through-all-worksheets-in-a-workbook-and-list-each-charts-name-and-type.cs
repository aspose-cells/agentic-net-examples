// Title: How to enumerate all charts in every worksheet of an Excel workbook and display their names and types using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# program that uses Aspose.Cells to open an .xlsx file, loops through each worksheet, and prints the worksheet name together with each chart's Name property and its ChartType enumeration value. | Provide sample code that demonstrates retrieving the Name and Type of every chart on all sheets of a workbook with Aspose.Cells, including handling missing files and runtime exceptions.
// Common Searches: Aspose.Cells C# list chart names and chart types for all worksheets | How to get chart type from each chart in an Excel file using Aspose.Cells .NET | Iterate through worksheets and retrieve chart information with Aspose.Cells | C# code example to enumerate charts in a workbook using Aspose.Cells | Display worksheet name, chart name, and chart type using Aspose.Cells
// Tags: enumerate charts Aspose.Cells .NET | retrieve chart name and type C# | loop worksheets Aspose.Cells chart extraction | Aspose.Cells chart enumeration example | Excel workbook chart metadata extraction

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook with Aspose.Cells, iterates over every worksheet, and for each chart prints the worksheet name, the chart's Name, and its ChartType, while safely handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts on the current worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Output the worksheet name, chart name, and chart type
                    Console.WriteLine($"Worksheet: {sheet.Name}, Chart Name: {chart.Name}, Chart Type: {chart.Type}");
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
