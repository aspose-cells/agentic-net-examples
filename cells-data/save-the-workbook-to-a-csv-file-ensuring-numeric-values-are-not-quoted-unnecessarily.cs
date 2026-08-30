// Title: Save an Aspose.Cells workbook to CSV in C# without quoting numeric cells
// AI Prompts: Generate C# code that uses Aspose.Cells to export a workbook to a CSV file, applying TxtSaveOptions so that only necessary values are quoted. | Show how to configure TxtSaveOptions with QuoteType.Minimum, UTF‑8 encoding, and a comma delimiter for CSV output in Aspose.Cells.
// Common Searches: Aspose.Cells C# export to CSV avoid quotes around numbers | How to configure TxtSaveOptions for minimal quoting in CSV files | CSV save options separator comma UTF-8 Aspose.Cells example | Prevent numeric values from being enclosed in quotes when saving Excel as CSV with Aspose.Cells
// Tags: Aspose.Cells CSV export minimal quoting | TxtSaveOptions QuoteType Minimum C# | CSV separator comma UTF-8 Aspose.Cells | prevent numeric quoting Aspose.Cells | save workbook as CSV Aspose.Cells C#

using System.Text;
using Aspose.Cells;

// Creates a workbook, adds sample data, sets TxtSaveOptions to use a comma delimiter, UTF‑8 encoding, and QuoteType.Minimum so numeric cells are written without surrounding quotes, then saves the file as output.csv.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("John");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["A3"].PutValue("Jane");
        worksheet.Cells["B3"].PutValue(25);

        // Configure CSV save options to avoid unnecessary quoting of numeric values
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
        {
            Separator = ',',               // Use comma as delimiter
            Encoding = Encoding.UTF8,      // UTF-8 encoding
            QuoteType = TxtValueQuoteType.Minimum // Quote only when truly needed
        };

        // Save the workbook as CSV
        workbook.Save("output.csv", saveOptions);
    }
}
