// Title: Create a C# audit log of QuotePrefix custom document property changes across sequential Excel workbooks with Aspose.Cells
// Common Searches: how to log changes of a custom document property like QuotePrefix in multiple Excel files using Aspose.Cells C# | compare QuotePrefix values between different workbook versions with Aspose.Cells for .NET | generate audit trail for Excel custom properties across versioned files in C# | detect and report modifications of QuotePrefix custom property in a series of .xlsx workbooks | Aspose.Cells example for tracking custom document property changes over time
// Tags: read custom document property Aspose.Cells C# | compare workbook custom properties Aspose.Cells | audit log Excel property changes C# | track QuotePrefix across workbook versions | sequential workbook processing Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample iterates through a chronologically ordered list of Excel files, loads each workbook with Aspose.Cells, reads the "QuotePrefix" custom document property (treating missing properties as empty), compares it with the previous version, and writes a console message whenever the value changes, while handling missing files and runtime exceptions.
class QuotePrefixAudit
{
    static void Main()
    {
        // List of workbook file paths in chronological order
        var workbookFiles = new List<string>
        {
            "Version1.xlsx",
            "Version2.xlsx",
            "Version3.xlsx"
        };

        string previousQuotePrefix = null;
        string previousFile = null;

        foreach (var file in workbookFiles)
        {
            try
            {
                // Ensure the file exists before attempting to load it
                if (!File.Exists(file))
                {
                    Console.WriteLine($"File not found: {file}");
                    continue;
                }

                // Load workbook
                var workbook = new Workbook(file);

                // Retrieve the current QuotePrefix setting from a custom document property
                // If the property does not exist, treat it as an empty string
                string currentQuotePrefix = string.Empty;
                if (workbook.CustomDocumentProperties["QuotePrefix"] != null)
                {
                    var prop = workbook.CustomDocumentProperties["QuotePrefix"];
                    currentQuotePrefix = prop.Value?.ToString() ?? string.Empty;
                }

                // Compare with the previous version, if any
                if (previousQuotePrefix != null && currentQuotePrefix != previousQuotePrefix)
                {
                    Console.WriteLine(
                        $"QuotePrefix changed from '{previousQuotePrefix}' in '{previousFile}' to '{currentQuotePrefix}' in '{file}'.");
                }

                // Update tracking variables for the next iteration
                previousQuotePrefix = currentQuotePrefix;
                previousFile = file;
            }
            catch (Exception ex)
            {
                // Log any unexpected errors and continue with the next file
                Console.WriteLine($"Error processing file '{file}': {ex.Message}");
            }
        }
    }
}
