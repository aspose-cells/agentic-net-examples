// Title: Create a case‑insensitive header‑to‑column index lookup dictionary by iterating the first row with Aspose.Cells for .NET (C#)
// Common Searches: how to build a header lookup dictionary from the first row in an Excel file using Aspose.Cells C# | C# code to get column index by header name with case‑insensitive matching in Aspose.Cells | iterate over Excel header row to map column names to indexes using Aspose.Cells .NET | load password protected workbook and create header‑to‑column map with Aspose.Cells | retrieve column number for a specific header value in a worksheet using Aspose.Cells C#
// Tags: header row iteration Aspose.Cells | case-insensitive column index map C# | Excel header to column dictionary .NET | load password protected workbook Aspose.Cells | trimmed header text mapping Aspose.Cells | save workbook after header processing C#

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook (optionally prompting for a password), accesses the first worksheet, iterates the header row to build a case‑insensitive Dictionary<string,int> that maps each trimmed header text to its column index, demonstrates retrieving the index for a specific header, and saves the workbook.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook = null;

        // First attempt: load without a password
        try
        {
            workbook = new Workbook(inputPath);
        }
        catch (CellsException)
        {
            // Assume the workbook is password protected
            Console.WriteLine("The workbook is password protected. Please provide a password.");
            Console.Write("Password: ");
            string pwd = Console.ReadLine() ?? string.Empty;

            var loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = pwd };
            try
            {
                workbook = new Workbook(inputPath, loadOptions);
            }
            catch (Exception innerEx)
            {
                Console.WriteLine($"Failed to load workbook with the provided password: {innerEx.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Header row (0‑based index)
        int headerRowIndex = 0;

        // Map header text to column index
        var headerLookup = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        int maxColumn = cells.MaxColumn;

        // Build the header lookup dictionary
        for (int col = 0; col <= maxColumn; col++)
        {
            Cell headerCell = cells[headerRowIndex, col];
            string headerText = headerCell?.StringValue?.Trim();

            if (!string.IsNullOrEmpty(headerText) && !headerLookup.ContainsKey(headerText))
            {
                headerLookup[headerText] = col;
            }
        }

        // Example: retrieve column index for "Name"
        if (headerLookup.TryGetValue("Name", out int nameColumnIndex))
        {
            Console.WriteLine($"Column index for 'Name' is {nameColumnIndex}");
        }

        // Save the workbook (optional)
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
