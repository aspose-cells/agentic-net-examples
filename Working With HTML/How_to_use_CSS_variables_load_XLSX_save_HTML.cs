using System;
using Aspose.Cells;

namespace AsposeCellsCssVariablesDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook from disk
            // (Replace "input.xlsx" with the actual path to your source file)
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable CSS custom properties (CSS variables) in the generated HTML
            // This makes Aspose.Cells emit reusable CSS variables for repeated values
            // such as colors, fonts, or base64 images.
            htmlOptions.EnableCssCustomProperties = true;

            // Optional: you can also set a prefix for generated CSS class names
            // htmlOptions.CellCssPrefix = "mycell-";

            // Save the workbook as an HTML file with CSS variables enabled
            // (Replace "output.html" with the desired output path)
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook has been converted to HTML with CSS variables.");
        }
    }
}