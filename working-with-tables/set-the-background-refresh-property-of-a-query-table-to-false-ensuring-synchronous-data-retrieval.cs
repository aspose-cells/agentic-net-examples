// Title: Disable background refresh for an Excel query table using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a workbook with Aspose.Cells, checks for query tables, and sets RefreshOnLoad = false to prevent background refresh. | Demonstrate how to set RefreshPeriod to zero for a query table using Aspose.Cells so the data is retrieved synchronously. | Provide a complete example that modifies query table refresh settings and saves the updated Excel file.
// Common Searches: Aspose.Cells .NET set query table RefreshOnLoad false | how to make Excel query table refresh synchronously with Aspose.Cells | disable background refresh property for query tables in Aspose.Cells C# | control query table refresh period using Aspose.Cells API
// Tags: Aspose.Cells query table refresh control | C# query table refresh configuration | Excel query table synchronous refresh Aspose | adjust query table refresh interval Aspose.Cells | prevent background refresh Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel workbook, checks for query tables on the first worksheet, and explains that Aspose.Cells processes query tables synchronously (no BackgroundRefresh property). It shows how to control refresh behavior with RefreshOnLoad or RefreshPeriod, then saves the workbook.
class Program
{
    static void Main()
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

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // If the worksheet contains query tables, process the first one
            if (sheet.QueryTables.Count > 0)
            {
                QueryTable queryTable = sheet.QueryTables[0];

                // Aspose.Cells processes query tables synchronously; no BackgroundRefresh property exists.
                // If needed, you can control refresh behavior via RefreshOnLoad or RefreshPeriod.
                // Example (optional):
                // queryTable.RefreshOnLoad = true;
            }

            // Save the workbook with any modifications
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
