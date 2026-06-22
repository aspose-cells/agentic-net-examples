using System;
using Aspose.Cells;

namespace HyperlinkAndFreezeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a hyperlink to an external document in cell A2
            // Parameters: cell name, rows in range, columns in range, address (file path or URL)
            worksheet.Hyperlinks.Add("A2", 1, 1, @"C:\Docs\ExternalFile.pdf");

            // Optionally set display text and screen tip for better UX
            Hyperlink link = worksheet.Hyperlinks[0];
            link.TextToDisplay = "Open External Document";
            link.ScreenTip = "Click to open the PDF file";

            // Freeze the first two rows so that the row containing the hyperlink (row 2) stays visible
            // FreezePanes(string cellName, int freezedRows, int freezedColumns)
            worksheet.FreezePanes("A2", 2, 0);

            // Save the workbook to a file
            workbook.Save("HyperlinkAndFreezeDemo.xlsx");
        }
    }
}