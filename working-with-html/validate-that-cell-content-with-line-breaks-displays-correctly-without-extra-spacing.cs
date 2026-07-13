using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Author: Aspose.Cells .NET example – validate line‑break rendering without extra spaces

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set an HTML string that contains <br> tags to create line breaks in cell A1
        // The HtmlString property interprets the HTML and renders line breaks correctly.
        worksheet.Cells["A1"].HtmlString = "First line<br>Second line<br>Third line";

        // Save the workbook – the cell will display the three lines without additional spacing.
        workbook.Save("CellLineBreaksValidated.xlsx", SaveFormat.Xlsx);
    }
}