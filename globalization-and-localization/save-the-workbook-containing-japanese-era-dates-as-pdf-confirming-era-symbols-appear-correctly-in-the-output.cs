using System;
using Aspose.Cells;

class JapaneseEraPdfDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Set the workbook's regional settings to Japan (required for era formatting)
        wb.Settings.Region = CountryCode.Japan;

        // Insert a serial date value (e.g., 44089 corresponds to 2020-09-15)
        Cell cell = wb.Worksheets[0].Cells[0, 0];
        cell.PutValue(44089);

        // Apply a custom format that displays the Japanese era (e.g., "令和2年9月15日")
        Style style = cell.GetStyle();
        style.Custom = "[$-F800]yyyy年m月d日";
        cell.SetStyle(style);

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Use a Japanese font to ensure era symbols render correctly
        pdfOptions.DefaultFont = "MS Gothic";
        // Try to use the workbook's default font first
        pdfOptions.CheckWorkbookDefaultFont = true;

        // Save the workbook as PDF
        wb.Save("JapaneseEra.pdf", pdfOptions);
    }
}