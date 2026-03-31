using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace HtmlToXlsxConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Desired path for the resulting XLSX workbook
            string xlsxPath = "output.xlsx";

            // Create HTML load options to control how the HTML is imported
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            loadOptions.LoadFormulas = true;      // Preserve formulas if present in the HTML
            loadOptions.SupportDivTag = true;    // Enable support for <div> layout tags
            loadOptions.AutoFitColsAndRows = true; // Adjust column widths and row heights automatically

            // Load the HTML file into a Workbook using the specified options
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Save the workbook as an XLSX file, preserving cell data and formatting
            workbook.Save(xlsxPath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML file '{htmlPath}' has been successfully converted to '{xlsxPath}'.");
        }
    }
}