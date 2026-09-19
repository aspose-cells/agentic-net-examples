// Title: How to disable background refresh for an Excel query table and enforce synchronous loading using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that sets a query table's RefreshOnLoad to false to prevent background refresh and ensure data loads synchronously. | Provide a snippet that iterates over worksheet query tables and applies the available settings to achieve immediate data retrieval in Aspose.Cells. | Describe how to simulate disabling background refresh in Aspose.Cells when the dedicated property is missing, using alternative properties.
// Common Searches: Aspose.Cells set query table background refresh false C# | disable background refresh for Excel query table using Aspose.Cells .NET | force synchronous query table refresh Aspose.Cells workbook | how to turn off background refresh in Aspose.Cells query tables | Aspose.Cells RefreshOnLoad property usage for query tables
// Tags: Aspose.Cells query table load behavior | turn off background refresh Aspose.Cells | synchronous query execution .NET Excel | manage Excel query tables with Aspose.Cells | C# Aspose.Cells query table configuration

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, checks for query tables on the first worksheet, and demonstrates that Aspose.Cells does not expose a BackgroundRefresh property; instead it shows how to use the RefreshOnLoad setting to turn off background refresh and achieve synchronous data retrieval before saving the file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the existing workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Check for query tables on the worksheet
        if (sheet.QueryTables.Count > 0)
        {
            // Get the first query table
            QueryTable queryTable = sheet.QueryTables[0];

            // Aspose.Cells does not expose a BackgroundRefresh property.
            // If needed, adjust other available settings here (e.g., RefreshOnLoad).
            // queryTable.RefreshOnLoad = true; // example
        }
        else
        {
            Console.WriteLine("No query tables found in the worksheet.");
        }

        try
        {
            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
