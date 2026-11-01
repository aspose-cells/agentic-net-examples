// Title: Enable background refresh for all data connections in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set background refresh for all data connections | How to make Excel data connections refresh asynchronously with Aspose.Cells | Enable background refresh property on workbook connections using Aspose.Cells .NET | Aspose.Cells example for improving UI responsiveness during data refresh | Set BackgroundRefresh = true for each DataConnection in a workbook with Aspose.Cells
// Tags: background refresh data connections Aspose.Cells | asynchronous workbook data refresh .NET | set BackgroundRefresh property C# | Excel data connection configuration Aspose.Cells | UI responsiveness during Excel refresh Aspose.Cells | modify workbook connections programmatically Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file, iterates through every DataConnection in the workbook, sets each connection's BackgroundRefresh flag to true for asynchronous updates, ensures the output folder exists, and saves the modified workbook, handling missing files and exceptions.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Enable background refresh for all data connections in the workbook
            foreach (var connection in workbook.DataConnections)
            {
                // When set to true, the connection refreshes data in the background,
                // allowing the UI to remain responsive during the operation.
                connection.BackgroundRefresh = true;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
