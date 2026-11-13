// Title: Rename the first data connection in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# rename first workbook data connection to custom name | set data connection name in Excel file with Aspose.Cells .NET | change Excel data connection identifier programmatically using Aspose.Cells | C# update DBConnection.Name property in Aspose.Cells workbook | how to modify data connection name in an existing .xlsx using Aspose.Cells
// Tags: rename workbook data connection Aspose.Cells | set data connection name C# | modify Excel data connections programmatically | Aspose.Cells data connection API | update DBConnection.Name .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel workbook, checks for any data connections, renames the first connection to "SalesDataConnection", ensures the output directory exists, and saves the modified workbook, while handling errors and reporting status to the console.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Rename the first data connection if any exist
            if (workbook.DataConnections != null && workbook.DataConnections.Count > 0)
            {
                var conn = workbook.DataConnections[0]; // Use var to avoid explicit type issues
                conn.Name = "SalesDataConnection";
                Console.WriteLine($"Renamed data connection to \"{conn.Name}\".");
            }
            else
            {
                Console.WriteLine("No data connections found in the workbook.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the updated connection name
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
