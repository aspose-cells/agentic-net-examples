// Title: C# – Export Aspose.Cells Workbook to CSV with Double‑Quote Text Qualifiers
// Description: Learn how to save an Aspose.Cells workbook as a CSV file in C#, using a comma delimiter, UTF‑8 encoding, and always enclosing each field in double quotes to preserve commas and other special characters. Ideal for global data exchange (US, UK, India, etc.).
// Keywords: Aspose.Cells CSV export C# | double‑quote text qualifier | TxtSaveOptions QuoteType Always | comma delimiter CSV | UTF‑8 CSV output | handle commas in CSV fields | global data export
// Common Searches: Aspose.Cells export CSV with quotes C# | How to always quote fields when saving CSV using Aspose.Cells | C# CSV export with double‑quote qualifier for every column | Save workbook as CSV with comma separator and quoted values
// Developer Intent: Generate a CSV file where every cell value is wrapped in double quotes to ensure correct parsing of commas and other delimiters.
// Use Cases: Creating CSV reports that include address or description fields containing commas. | Producing data feeds for third‑party systems that require all values to be quoted. | Exporting multilingual datasets with UTF‑8 encoding and consistent text qualifiers.
// AI Prompts: Provide C# code that uses Aspose.Cells to export a workbook to CSV with every field enclosed in double quotes. | Show how to configure TxtSaveOptions for CSV: comma separator, UTF‑8 encoding, and QuoteType set to Always. | Demonstrate saving a workbook to a specific file path as a quoted‑field CSV using Aspose.Cells.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    // Learn how to save an Aspose.Cells workbook as a CSV file in C#, using a comma delimiter, UTF‑8 encoding, and always enclosing each field in double quotes to preserve commas and other special characters. Ideal for global data exchange (US, UK, India, etc.).
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate cells with data that contains commas
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Address");
            cells["A2"].PutValue("John Doe");
            cells["B2"].PutValue("123 Main St, Springfield");
            cells["A3"].PutValue("Jane Smith");
            cells["B3"].PutValue("456 Oak Ave, Metropolis");

            // Configure CSV save options
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Separator = ',',                 // Use comma as delimiter
                Encoding = Encoding.UTF8,        // UTF-8 encoding
                QuoteType = TxtValueQuoteType.Always // Enclose every field in double quotes
            };

            // Save the workbook as CSV with the specified options
            string outputPath = "ExportedData.csv";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook exported to CSV with double‑quote qualifiers: {outputPath}");
        }
    }
}
