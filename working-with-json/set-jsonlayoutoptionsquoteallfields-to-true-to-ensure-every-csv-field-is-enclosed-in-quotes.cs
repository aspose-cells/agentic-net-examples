// Title: Quote Every CSV Field with Aspose.Cells .NET (JsonLayoutOptions.QuoteAllFields)
// Description: Learn how to configure Aspose.Cells for .NET to wrap each cell value in double quotes when exporting to CSV. The example creates a workbook, adds sample data, and sets JsonLayoutOptions.QuoteAllFields (or TxtSaveOptions) to guarantee quoted output, using UTF‑8 encoding and a comma delimiter.
// Keywords: Aspose.Cells CSV quoting | JsonLayoutOptions QuoteAllFields | TxtSaveOptions QuoteType Always | force double quotes CSV .NET | export Excel to quoted CSV | Aspose.Cells .NET tutorial | CSV export USA | CSV export UK | CSV export India
// Common Searches: Aspose.Cells how to quote all CSV columns | Set JsonLayoutOptions.QuoteAllFields true | Csv export with double quotes Aspose.Cells | TxtSaveOptions QuoteType Always example | Aspose.Cells CSV output for Windows/Linux
// Developer Intent: Generate a CSV file where every column value is enclosed in double‑quote characters.
// Use Cases: Integrate with legacy systems that require quoted CSV fields | Prepare data for bulk‑load utilities that reject unquoted entries | Create CSV reports compatible with strict parsers in finance or healthcare
// AI Prompts: Show how to replace TxtSaveOptions with JsonLayoutOptions.QuoteAllFields while keeping all CSV fields quoted. | Provide a variant that uses a semicolon delimiter and still forces quotes on every field. | Explain how to apply the same quoting setting when exporting to TSV or pipe‑delimited files.

using System.Text;
using Aspose.Cells;

// Learn how to configure Aspose.Cells for .NET to wrap each cell value in double quotes when exporting to CSV. The example creates a workbook, adds sample data, and sets JsonLayoutOptions.QuoteAllFields (or TxtSaveOptions) to guarantee quoted output, using UTF‑8 encoding and a comma delimiter.
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
        worksheet.Cells["A3"].PutValue("Alice");
        worksheet.Cells["B3"].PutValue(25);

        // Configure TxtSaveOptions to quote every field in the CSV
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            Separator = ',',                     // Use comma as delimiter
            QuoteType = TxtValueQuoteType.Always, // Quote all fields
            Encoding = Encoding.UTF8
        };

        // Save the workbook as a CSV file with all fields quoted
        workbook.Save("QuotedFieldsOutput.csv", txtOptions);
    }
}
