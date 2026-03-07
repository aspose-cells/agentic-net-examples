using System;
using Aspose.Cells;

namespace AsposeCellsHtmlToPdf
{
    public class Program
    {
        public static void Main()
        {
            // Load the HTML file into the workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook("input.html", loadOptions);

            // Save the workbook as a PDF file
            workbook.Save("output.pdf", SaveFormat.Pdf);

            Console.WriteLine("HTML file successfully converted to PDF.");
        }
    }
}