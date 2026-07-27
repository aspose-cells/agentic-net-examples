using System.Text;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Alice");
        sheet.Cells["B3"].PutValue(25);

        // Configure TxtSaveOptions so that every field is enclosed in quotes
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            Separator = ',',                     // Use comma as CSV separator
            QuoteType = TxtValueQuoteType.Always, // Quote every field
            Encoding = Encoding.UTF8
        };

        // Save the workbook as a CSV file with all fields quoted
        workbook.Save("AllFieldsQuoted.csv", txtOptions);
    }
}