using System;
using System.Text;
using Aspose.Cells;

class ExportWorksheetToCsvAsText
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate cells with different data types
        worksheet.Cells["A1"].PutValue("Date");
        worksheet.Cells["B1"].PutValue(DateTime.Now);               // DateTime value
        worksheet.Cells["A2"].PutValue("Number");
        worksheet.Cells["B2"].PutValue(12345.678);                  // Numeric value
        worksheet.Cells["A3"].PutValue("Text");
        worksheet.Cells["B3"].PutValue("Sample text");             // String value

        // Apply a number format to the numeric cell so that its displayed text is formatted
        Style numberStyle = workbook.CreateStyle();
        numberStyle.Number = 2; // two decimal places
        worksheet.Cells["B2"].SetStyle(numberStyle);

        // Configure TxtSaveOptions to export cell values as their displayed text (preserving types as text)
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
        {
            // Use the displayed style of each cell when converting to string
            FormatStrategy = CellValueFormatStrategy.DisplayStyle,
            // Quote all fields to ensure correct CSV parsing
            QuoteType = TxtValueQuoteType.Always,
            // Use UTF-8 encoding
            Encoding = Encoding.UTF8,
            // Export only the active sheet (default behavior)
            ExportAllSheets = false
        };

        // Save the worksheet as a CSV file with the specified options
        workbook.Save("output.csv", saveOptions);
    }
}