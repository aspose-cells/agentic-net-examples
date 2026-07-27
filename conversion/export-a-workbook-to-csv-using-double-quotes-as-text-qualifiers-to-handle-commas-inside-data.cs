// Title: Export Aspose.Cells Workbook to CSV with Double‑Quote Qualifiers (C#)
// Description: Shows how to create a workbook, add cells that contain commas, and save it as a CSV using Aspose.Cells TxtSaveOptions with a comma separator, UTF‑8 encoding, and QuoteType.Always so every field is wrapped in double quotes.
// Keywords: Aspose.Cells CSV export C# | TxtSaveOptions QuoteType.Always | double quote text qualifier | comma delimiter CSV | C# export Excel to CSV | UTF‑8 CSV Aspose | handle commas in CSV | Aspose.Cells .NET | CSV quoting options | SaveFormat.Csv
// Common Searches: Aspose.Cells export CSV with all fields quoted | C# TxtSaveOptions QuoteType.Always example | force double quotes in CSV using Aspose.Cells | CSV export commas inside data Aspose .NET | save Excel as CSV with text qualifier Aspose
// Developer Intent: Generate a CSV file from an Aspose.Cells workbook where each value is enclosed in double quotes to safely preserve commas and other delimiters.
// Use Cases: Export address columns that contain commas for CRM imports that require quoted CSV. | Create CSV files for legacy parsers that expect every field to be quoted. | Produce locale‑independent CSV for international data exchange pipelines. | Automate batch conversion of Excel reports to CSV for downstream ETL processes.
// AI Prompts: Write C# code using Aspose.Cells to export a workbook to CSV with QuoteType.Always and UTF‑8 encoding. | Explain why QuoteType.Always is needed when cell values contain commas in a CSV export. | Show how to configure TxtSaveOptions with a comma separator, UTF‑8 encoding, and always‑quote fields for CSV output.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    // Shows how to create a workbook, add cells that contain commas, and save it as a CSV using Aspose.Cells TxtSaveOptions with a comma separator, UTF‑8 encoding, and QuoteType.Always so every field is wrapped in double quotes.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate cells with data that includes commas
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Address");
            cells["A2"].PutValue("John Doe");
            cells["B2"].PutValue("123 Main St, Springfield");
            cells["A3"].PutValue("Jane Smith");
            cells["B3"].PutValue("456 Oak Ave, Metropolis");

            // Configure CSV save options:
            // - Use comma as separator
            // - Quote all fields with double quotes to safely handle commas inside data
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Separator = ',',
                QuoteType = TxtValueQuoteType.Always,
                Encoding = Encoding.UTF8
            };

            // Save the workbook as CSV with the specified options
            workbook.Save("ExportedData.csv", saveOptions);
        }
    }
}
