using System;
using Aspose.Cells;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put placeholder text into the target cell
        worksheet.Cells["A1"].PutValue("Send Email");

        // Add a mailto hyperlink to cell A1
        // Parameters: cell name, rows in range, columns in range, hyperlink address
        int hyperlinkIndex = worksheet.Hyperlinks.Add("A1", 1, 1, "mailto:john.doe@example.com");

        // Optionally change the displayed text of the hyperlink
        worksheet.Hyperlinks[hyperlinkIndex].TextToDisplay = "Email John Doe";

        // Apply typical hyperlink styling (blue and underlined)
        Style style = worksheet.Cells["A1"].GetStyle();
        style.Font.Color = Color.Blue;
        style.Font.Underline = FontUnderlineType.Single;
        worksheet.Cells["A1"].SetStyle(style);

        // Save the workbook
        workbook.Save("MailtoHyperlink.xlsx");
    }
}