using System;
using Aspose.Cells;

namespace AsposeCellsPreserveDivLayoutExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source HTML file that contains <div> tags
            string htmlPath = "input.html";

            // Path for the resulting Excel file
            string excelPath = "output.xlsx";

            // Create HtmlLoadOptions and enable support for <div> tag layout
            HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
            loadOptions.SupportDivTag = true; // PreserveDivLayout equivalent

            // Load the HTML file into a Workbook using the configured options
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Save the workbook to XLSX format
            workbook.Save(excelPath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML file '{htmlPath}' has been converted to Excel file '{excelPath}' with PreserveDivLayout enabled.");
        }
    }
}