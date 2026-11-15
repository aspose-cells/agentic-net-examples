// Title: Read a password‑protected XLS workbook and extract the first chart from the second worksheet using Aspose.Cells for .NET
// Common Searches: how to load a password protected .xls file with Aspose.Cells in C# | retrieve chart from the second worksheet of an encrypted Excel workbook using Aspose.Cells | Aspose.Cells example for opening protected workbook and accessing its charts | C# code to get chart name and type from a password‑protected XLS file | using LoadOptions to read a protected Excel file and extract chart information
// Tags: load password protected XLS with LoadOptions | second worksheet chart extraction Aspose.Cells | first chart retrieval from sheet C# | encrypted workbook chart data access | Aspose.Cells chart extraction from protected file

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample verifies the XLS file exists, loads the password‑protected workbook via LoadOptions, selects the second worksheet (index 1), checks for charts, retrieves the first chart, and prints its name, type, and series count, with proper exception handling.
class Program
{
    static void Main()
    {
        // Path to the password‑protected XLS file
        string filePath = @"C:\Path\To\Your\File.xls";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        // Password for the protected workbook
        string password = "yourPassword";

        try
        {
            // Load the workbook with the password using LoadOptions (format auto‑detected)
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Access the second worksheet (index 1 because collection is zero‑based)
            Worksheet secondSheet = workbook.Worksheets[1];

            // Ensure the worksheet contains at least one chart
            if (secondSheet.Charts.Count > 0)
            {
                // Retrieve the first chart on the second worksheet
                Chart chart = secondSheet.Charts[0];

                // Output basic chart information
                Console.WriteLine($"Chart Name: {chart.Name}");
                Console.WriteLine($"Chart Type: {chart.Type}");
                Console.WriteLine($"Number of Series: {chart.NSeries.Count}");
            }
            else
            {
                Console.WriteLine("No charts found on the second worksheet.");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
