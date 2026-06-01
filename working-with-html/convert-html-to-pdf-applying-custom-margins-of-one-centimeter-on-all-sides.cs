using System;
using Aspose.Cells;

namespace HtmlToPdfWithMargins
{
    class Program
    {
        static void Main()
        {
            // Load the HTML file into a workbook
            // Replace "input.html" with the path to your source HTML file
            Workbook workbook = new Workbook("input.html");

            // Apply 1 centimeter margins to the first worksheet (you can repeat for other sheets if needed)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.TopMargin = 1.0;      // Top margin in centimeters
            sheet.PageSetup.BottomMargin = 1.0;   // Bottom margin in centimeters
            sheet.PageSetup.LeftMargin = 1.0;     // Left margin in centimeters
            sheet.PageSetup.RightMargin = 1.0;    // Right margin in centimeters

            // Save the workbook as a PDF file
            // Replace "output.pdf" with the desired output path
            workbook.Save("output.pdf", SaveFormat.Pdf);

            Console.WriteLine("HTML has been converted to PDF with 1 cm margins.");
        }
    }
}