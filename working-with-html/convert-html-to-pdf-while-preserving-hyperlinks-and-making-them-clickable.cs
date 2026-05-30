using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsHtmlToPdf
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Path for the resulting PDF file
            string pdfPath = "output.pdf";

            // Load the HTML file into a Workbook.
            // LoadOptions with LoadFormat.Html ensures correct parsing of hyperlinks.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Save the workbook as PDF. Hyperlinks are preserved and become clickable by default.
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine("HTML has been converted to PDF with clickable hyperlinks.");
        }
    }
}