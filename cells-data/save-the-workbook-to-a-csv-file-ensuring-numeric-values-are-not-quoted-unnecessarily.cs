using System.Text;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the worksheet with sample data
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);   // numeric value
        sheet.Cells["A3"].PutValue("Jane");
        sheet.Cells["B3"].PutValue(25);   // numeric value

        // Configure CSV save options:
        // - Use comma as separator (default for CSV)
        // - UTF-8 encoding
        // - QuoteType.Normal quotes only when necessary, so numbers are not quoted
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
        {
            Separator = ',',
            Encoding = Encoding.UTF8,
            QuoteType = TxtValueQuoteType.Normal
        };

        // Save the workbook as a CSV file with the specified options
        workbook.Save("output.csv", csvOptions);
    }
}