// Title: Export Cells with Leading Apostrophes to CSV using Aspose.Cells for .NET
// Description: C# example that creates a workbook, writes values beginning with an apostrophe, sets TxtSaveOptions (UTF‑8, comma separator, ExportQuotePrefix = true, Minimum quoting) and saves the sheet as a CSV file where the leading apostrophes are kept as literal characters.
// Keywords: Aspose.Cells CSV export | ExportQuotePrefix | leading apostrophe CSV | C# Aspose.Cells TxtSaveOptions | UTF-8 CSV Aspose
// Common Searches: Aspose.Cells keep leading apostrophe when saving CSV | ExportQuotePrefix true C# example | save Excel as CSV with apostrophe character | minimal quoting CSV Aspose.Cells | UTF-8 CSV export Aspose.Cells .NET
// Developer Intent: Save a worksheet to CSV while preserving any leading apostrophe characters in cell values.
// Use Cases: Export product codes or identifiers that start with an apostrophe without losing the character. | Generate CSV files for legacy systems that require literal apostrophes in text fields. | Create UTF‑8 encoded CSV reports with minimal quoting while retaining formatting markers.
// AI Prompts: Show how to also preserve embedded double quotes when exporting to CSV with Aspose.Cells. | Provide a variant that reads an existing Excel file, applies ExportQuotePrefix, and writes a semicolon‑delimited CSV. | Explain how ExportQuotePrefix interacts with the QuoteType setting in TxtSaveOptions.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// C# example that creates a workbook, writes values beginning with an apostrophe, sets TxtSaveOptions (UTF‑8, comma separator, ExportQuotePrefix = true, Minimum quoting) and saves the sheet as a CSV file where the leading apostrophes are kept as literal characters.
class ExportRichTextCsv
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Example 1: Text with leading apostrophe to force text format
            // The apostrophe will be exported as a literal character because ExportQuotePrefix is true
            cells["A1"].PutValue("'Hello World");

            // Example 2: Plain text with leading apostrophe
            cells["B1"].PutValue("'Sample");

            // Configure CSV save options
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Separator = ',',                 // CSV delimiter
                Encoding = Encoding.UTF8,        // Use UTF‑8 encoding
                ExportQuotePrefix = true,        // Export the leading apostrophe as a literal character
                QuoteType = TxtValueQuoteType.Minimum // Quote only when necessary
            };

            // Define output file path
            string outputPath = "RichTextOutput.csv";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a CSV file
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
