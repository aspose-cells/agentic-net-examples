// Title: Sequentially refresh all external data connections in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file using Aspose.Cells, enumerates each DataConnection, and calls its Refresh method only if the method exists, handling any per‑connection exceptions. | Show how to use C# reflection to safely invoke the Refresh method on Aspose.Cells DataConnection objects, then save the updated workbook to a new file.
// Common Searches: Aspose.Cells C# refresh each external data connection in a workbook | how to loop through DataConnections and call Refresh with Aspose.Cells | use reflection to call Refresh on Excel data connections in .NET | save workbook after refreshing external connections using Aspose.Cells | handle missing Refresh method on Aspose.Cells data connections
// Tags: refresh data connections Aspose.Cells | enumerate DataConnections C# | reflection invoke Refresh Aspose.Cells | save workbook after connection refresh | fallback when Refresh method unavailable

using Aspose.Cells;
using System;
using System.IO;
using System.Reflection;

// The example loads an Excel workbook with Aspose.Cells, iterates over all DataConnection objects, uses reflection to invoke each connection's Refresh method only when it exists, logs any errors, ensures the output directory is present, and saves the refreshed workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Refresh external data connections if the Refresh method is available
            try
            {
                foreach (var connection in workbook.DataConnections)
                {
                    try
                    {
                        // Use reflection to call Refresh only when it exists
                        MethodInfo refreshMethod = connection.GetType().GetMethod("Refresh", BindingFlags.Public | BindingFlags.Instance);
                        if (refreshMethod != null)
                        {
                            refreshMethod.Invoke(connection, null);
                        }
                    }
                    catch (Exception connEx)
                    {
                        Console.WriteLine($"Failed to refresh connection '{connection.Name}': {connEx.Message}");
                    }
                }
            }
            catch (Exception refreshEx)
            {
                Console.WriteLine($"Failed to process data connections: {refreshEx.Message}");
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
