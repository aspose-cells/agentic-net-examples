using System.Text;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
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
        sheet.Cells["A3"].PutValue("Alice");
        sheet.Cells["B3"].PutValue(25);

        // TxtSaveOptions is used for CSV/TXT export.
        // JsonLayoutOptions does not contain a QuoteAllFields property,
        // so we set QuoteType to Always to ensure every field is quoted.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            Separator = ',',                 // CSV separator
            QuoteType = TxtValueQuoteType.Always, // Quote every field
            Encoding = Encoding.UTF8
        };

        // Save the workbook as a CSV file with all fields quoted
        workbook.Save("output.csv", txtOptions);
    }
}