using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert HTML‑formatted rich text into cell A1
        // The HtmlString property preserves the HTML tags when the file is saved.
        worksheet.Cells["A1"].HtmlString = 
            "This is <b>bold</b> and <i>italic</i> text<br/>with a line break";

        // Save the workbook as an XLSX file
        workbook.Save("RichText.xlsx", SaveFormat.Xlsx);
    }
}