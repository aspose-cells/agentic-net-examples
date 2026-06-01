using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class HyperlinkPdfDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put display text into cell A1
        sheet.Cells["A1"].PutValue("Visit Aspose");

        // Add a hyperlink to cell A1 (row 0, column 0, 1 row, 1 column)
        sheet.Hyperlinks.Add(0, 0, 1, 1, "https://www.aspose.com");

        // Define the output PDF file path
        string pdfPath = "HyperlinkDemo.pdf";

        // Save the workbook as PDF
        workbook.Save(pdfPath, SaveFormat.Pdf);

        // Simple verification: read the PDF file as text and check for the URL
        // (PDF stores URLs as plain text, so this works for basic validation)
        string pdfContent = File.ReadAllText(pdfPath, Encoding.Default);
        bool linkExists = pdfContent.Contains("https://www.aspose.com");

        Console.WriteLine("Hyperlink present in PDF: " + linkExists);
    }
}