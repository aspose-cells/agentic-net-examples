// Title: Save Aspose.Cells Workbook to CSV in C# Without Quoting Numbers
// Description: Demonstrates how to export a workbook created with Aspose.Cells to a CSV file using C#. The example sets TxtSaveOptions for CSV format with a comma separator, UTF‑8 encoding, normal quoting, and disables always‑quoted mode so numeric cells are written without surrounding quotes.
// Keywords: Aspose.Cells CSV export C# | TxtSaveOptions CSV | SaveFormat.Csv numeric values unquoted | UTF-8 CSV Aspose.Cells | .NET workbook to CSV | quote only when needed Aspose | CSV separator Aspose.Cells
// Common Searches: Aspose.Cells export to CSV without quotes | C# save workbook as CSV numeric values not quoted | TxtSaveOptions CSV separator UTF-8 | How to prevent quoting of numbers in Aspose.Cells CSV | Aspose.Cells CSV encoding and quoting options
// Developer Intent: Export a workbook to CSV while keeping numeric cells unquoted.
// Use Cases: Generate lightweight CSV reports for data pipelines that require plain numbers. | Create UTF‑8 encoded CSV files for database imports where quoted numbers cause parsing errors. | Automate CSV exports with custom separators and minimal quoting to reduce file size.
// AI Prompts: Write C# code using Aspose.Cells to save a workbook as CSV with a comma separator, UTF‑8 encoding, and quoting only text values. | Show how to configure TxtSaveOptions so numeric cells are not surrounded by quotes when exporting to CSV. | Explain how to change the CSV separator to a semicolon while still avoiding unnecessary quotes.

using System.Text;
using Aspose.Cells;

// Demonstrates how to export a workbook created with Aspose.Cells to a CSV file using C#. The example sets TxtSaveOptions for CSV format with a comma separator, UTF‑8 encoding, normal quoting, and disables always‑quoted mode so numeric cells are written without surrounding quotes.
class CsvExportExample
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Jane");
        sheet.Cells["B3"].PutValue(25);

        // Configure CSV save options: comma separator, UTF-8 encoding,
        // quote only when necessary (numeric values will not be quoted)
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
        {
            Separator = ',',
            Encoding = Encoding.UTF8,
            QuoteType = TxtValueQuoteType.Normal,
            AlwaysQuoted = false // obsolete property, kept for completeness
        };

        // Save the workbook as a CSV file using the configured options
        workbook.Save("output.csv", csvOptions);
    }
}
