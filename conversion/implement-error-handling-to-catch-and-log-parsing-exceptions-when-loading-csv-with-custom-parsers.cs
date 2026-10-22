// Title: Add try‑catch error handling for CSV loading with custom TxtLoadOptions in Aspose.Cells (.NET)
// Common Searches: C# how to catch exceptions when loading a CSV file with Aspose.Cells TxtLoadOptions | log parsing errors while converting semicolon‑separated CSV to XLSX using Aspose.Cells | Aspose.Cells .NET custom CSV separator error handling example | handle file not found and load errors for CSV to Excel conversion in C# | try‑catch around Workbook constructor for CSV load failures Aspose.Cells
// Tags: Aspose.Cells CSV load error handling | TxtLoadOptions custom separator exception | C# workbook load exception logging | CSV to XLSX conversion with try-catch | Aspose.Cells parsing exception logging

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The sample verifies the CSV file exists, configures TxtLoadOptions with a semicolon separator and UTF‑8 encoding, and wraps the Workbook constructor in a try‑catch block to capture and log any parsing exceptions. After optional processing, the workbook is saved to XLSX inside another try‑catch to log saving failures.
class Program
{
    static void Main()
    {
        const string inputPath = "input.csv";
        const string outputPath = "output.xlsx";

        // Verify that the input CSV file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.Error.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        // Configure CSV load options (separator and encoding) using TxtLoadOptions
        var loadOptions = new TxtLoadOptions(LoadFormat.CSV)
        {
            Separator = ';',          // custom column separator
            Encoding = Encoding.UTF8 // custom text encoding
        };

        Workbook workbook = null;

        // Load the CSV file using the configured options
        try
        {
            workbook = new Workbook(inputPath, loadOptions);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error loading CSV file: {ex.Message}");
            return;
        }

        // Process the workbook as needed
        // ...

        // Save the workbook to another format (e.g., XLSX) with error handling
        try
        {
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
