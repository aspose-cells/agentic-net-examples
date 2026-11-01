// Title: Load a password‑protected XLS workbook and read the first chart from the second worksheet using Aspose.Cells for .NET
// Common Searches: asp.net load password protected xls file with Aspose.Cells LoadOptions | c# get chart from second worksheet in a protected Excel workbook using Aspose.Cells | how to read chart properties from a password protected XLS in C# | retrieve chart series count from protected workbook Aspose.Cells | open encrypted Excel file and access charts with Aspose.Cells .NET
// Tags: load protected XLS using Aspose.Cells LoadOptions | retrieve first chart from worksheet index 1 | extract chart name and type C# | Aspose.Cells workbook password decryption | chart series count extraction Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example verifies the XLS file exists, loads the password‑protected workbook with LoadOptions, accesses the second worksheet (index 1), checks for charts, and if present outputs the first chart's name, type, and series count, while handling any exceptions.
class Program
{
    static void Main()
    {
        // Path to the password‑protected XLS file
        string filePath = @"C:\Path\To\Your\ProtectedFile.xls";

        // Verify that the file exists to prevent FileNotFoundException
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

            // Retrieve the first chart on the second worksheet (if any)
            if (secondSheet.Charts.Count > 0)
            {
                Chart chart = secondSheet.Charts[0];

                // Output some chart information
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
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
