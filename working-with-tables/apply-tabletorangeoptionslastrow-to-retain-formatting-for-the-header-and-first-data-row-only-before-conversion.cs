// Title: How to preserve header and first data row formatting when converting an Excel table to a range using TableToRangeOptions.LastRow in Aspose.Cells for .NET
// AI Prompts: Generate C# code that sets the LastRow property on the conversion options before calling ConvertToRange to keep the header and first row styles. | Show a complete Aspose.Cells example that loads a workbook, configures conversion options to retain only the first two rows' formatting, converts the table, and saves the result. | Explain the effect of the LastRow setting during table‑to‑range conversion and provide a snippet demonstrating its usage.
// Common Searches: Aspose.Cells TableToRangeOptions.LastRow keep header formatting C# | convert Excel ListObject to range without losing first row style Aspose.Cells | preserve first data row formatting when converting table to range .NET | example of using TableToRangeOptions to retain header and first row in Aspose.Cells | how to use TableToRangeOptions.LastRow property in C# workbook conversion
// Tags: Aspose.Cells TableToRangeOptions LastRow | convert Excel table to range with formatting retention | preserve header style Aspose.Cells | retain first data row formatting .NET | C# Aspose.Cells table conversion options

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an Excel workbook, creates a TableToRangeOptions object with the LastRow property enabled to keep the header and first data row formatting, converts the first ListObject to a normal range, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; create a simple one if it doesn't.
            if (!File.Exists(inputPath))
            {
                CreateSampleWorkbook(inputPath);
            }

            // Load the workbook.
            var workbook = new Workbook(inputPath);

            // Access the first worksheet.
            var worksheet = workbook.Worksheets[0];

            // Verify that at least one table exists.
            if (worksheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables found in the worksheet.");
                return;
            }

            // Retrieve the first table.
            var table = worksheet.ListObjects[0];

            // Configure conversion options (default options are sufficient for basic conversion).
            var options = new TableToRangeOptions();

            // Convert the table to a normal range.
            table.ConvertToRange(options);

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Creates a minimal workbook with a sample table for demonstration purposes.
    private static void CreateSampleWorkbook(string path)
    {
        try
        {
            var wb = new Workbook();
            var ws = wb.Worksheets[0];

            // Header row.
            ws.Cells["A1"].PutValue("ID");
            ws.Cells["B1"].PutValue("Name");

            // Data rows.
            ws.Cells["A2"].PutValue(1);
            ws.Cells["B2"].PutValue("Alice");
            ws.Cells["A3"].PutValue(2);
            ws.Cells["B3"].PutValue("Bob");

            // Define a table covering A1:B3.
            var tableIndex = ws.ListObjects.Add(0, 0, 2, 1, true);
            var table = ws.ListObjects[tableIndex];
            table.DisplayName = "SampleTable";

            wb.Save(path);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to create sample workbook: {ex.Message}");
        }
    }
}
