using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add an email hyperlink to cell N3 (mailto link)
        int hyperlinkIndex = worksheet.Hyperlinks.Add("N3", 1, 1, "mailto:john.doe@example.com");

        // Set the text that will be displayed in the cell
        worksheet.Hyperlinks[hyperlinkIndex].TextToDisplay = "Contact John Doe";

        // Save the workbook
        workbook.Save("EmailHyperlink.xlsx");
    }
}