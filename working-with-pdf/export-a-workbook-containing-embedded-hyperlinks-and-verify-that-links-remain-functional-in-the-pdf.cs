using System;
using System.IO;
using Aspose.Cells;

namespace HyperlinkPdfExportDemo
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a hyperlink to cell A1
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, address
            sheet.Hyperlinks.Add(0, 0, 1, 1, "https://www.aspose.com");

            // Optionally set display text and screen tip
            Hyperlink link = sheet.Hyperlinks[0];
            link.TextToDisplay = "Aspose Home";
            link.ScreenTip = "Open Aspose website";

            // Save the workbook as Excel (save rule)
            string excelPath = "HyperlinkDemo.xlsx";
            workbook.Save(excelPath);

            // Save the workbook as PDF (save rule)
            string pdfPath = "HyperlinkDemo.pdf";
            workbook.Save(pdfPath);

            // Verify that the PDF contains the hyperlink address
            // Simple verification by searching the raw PDF content for the URL string
            bool hyperlinkFound = false;
            if (File.Exists(pdfPath))
            {
                string pdfContent = File.ReadAllText(pdfPath);
                hyperlinkFound = pdfContent.Contains("https://www.aspose.com");
            }

            Console.WriteLine($"Hyperlink present in PDF: {hyperlinkFound}");
        }
    }
}