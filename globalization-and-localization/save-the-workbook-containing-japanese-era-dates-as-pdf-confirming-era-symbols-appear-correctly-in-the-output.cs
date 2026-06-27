using System;
using Aspose.Cells;

class JapaneseEraPdfDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Set the workbook's regional settings to Japan (enables Japanese era formatting)
        wb.Settings.Region = CountryCode.Japan;

        // Access the first worksheet
        Worksheet ws = wb.Worksheets[0];

        // Insert a sample date (e.g., 2020-09-15)
        Cell cell = ws.Cells["A1"];
        cell.PutValue(new DateTime(2020, 9, 15));

        // Apply a custom number format that displays the Japanese era
        // Format: era name (ggge) year month day, e.g., "令和2年9月15日"
        Style style = cell.GetStyle();
        style.Custom = "[$-ja-JP]ggge\"年\"m\"月\"d\"日\"";
        cell.SetStyle(style);

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Use a Japanese font that contains era symbols
        pdfOptions.DefaultFont = "MS Gothic";
        // Ensure the workbook's default font is considered (default is true)
        pdfOptions.CheckWorkbookDefaultFont = true;

        // Save the workbook as PDF
        wb.Save("JapaneseEra.pdf", pdfOptions);
    }
}