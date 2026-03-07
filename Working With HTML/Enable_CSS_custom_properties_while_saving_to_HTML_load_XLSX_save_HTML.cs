using System;
using Aspose.Cells;

namespace AsposeCellsHtmlCssCustomPropertiesDemo
{
    class Program
    {
        static void Main()
        {
            // Load the source XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable CSS custom properties to optimize the HTML output
            htmlOptions.EnableCssCustomProperties = true;

            // Save the workbook as HTML using the configured options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook saved to HTML with CSS custom properties enabled.");
        }
    }
}