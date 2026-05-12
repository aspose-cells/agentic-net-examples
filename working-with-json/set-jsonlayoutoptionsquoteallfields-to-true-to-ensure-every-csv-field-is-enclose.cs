using System.Text;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("John");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["A3"].PutValue("Alice");
        worksheet.Cells["B3"].PutValue(25);

        // TxtSaveOptions is used for CSV/TXT export.
        // Setting QuoteType to TxtValueQuoteType.Always forces every field to be enclosed in quotes.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            Separator = ',',               // Use comma as CSV separator
            QuoteType = TxtValueQuoteType.Always, // Quote all fields
            Encoding = Encoding.UTF8
        };

        // Save the workbook as a CSV file with all fields quoted
        workbook.Save("output.csv", saveOptions);
    }
}