using System;
using System.Text;
using Aspose.Cells;

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
        worksheet.Cells["B2"].PutValue(30);   // numeric value
        worksheet.Cells["A3"].PutValue("Jane");
        worksheet.Cells["B3"].PutValue(25);   // numeric value

        // Configure CSV save options:
        // - Use comma as separator
        // - Quote only when necessary (numeric values will not be quoted)
        TxtSaveOptions csvOptions = new TxtSaveOptions
        {
            Separator = ',',
            QuoteType = TxtValueQuoteType.Normal,
            Encoding = Encoding.UTF8
        };

        // Save the workbook as CSV using the configured options
        workbook.Save("output.csv", csvOptions);
    }
}