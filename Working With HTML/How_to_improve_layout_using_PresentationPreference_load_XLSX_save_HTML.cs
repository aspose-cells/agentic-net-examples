using System;
using Aspose.Cells;

namespace AsposeCellsPresentationPreferenceDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file (XLSX)
            string sourcePath = "input.xlsx";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable presentation preference for a more beautiful layout
            // This makes the generated HTML resemble the Excel presentation view
            htmlOptions.PresentationPreference = true;

            // Optional: set HTML version to HTML5 for modern browsers
            htmlOptions.HtmlVersion = HtmlVersion.Html5;

            // Save the workbook as an HTML file using the configured options
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook saved to HTML with PresentationPreference enabled at: {outputPath}");
        }
    }
}