// Title: Export an Excel workbook to HTML with French (fr-FR) number formatting using Aspose.Cells for .NET
// Common Searches: c# aspnet export excel to html with fr-fr number format using aspose.cells | set workbook cultureinfo before saving as html in aspose.cells | asp.net generate html from xlsx preserving locale specific numeric formatting | aspose.cells htmlsaveoptions exportgridlines french number formatting | convert excel to html with locale specific formatting asp.net core
// Tags: Aspose.Cells HTML export CultureInfo | C# set workbook culture for numeric formatting | HtmlSaveOptions ExportGridLines | locale-specific number formatting in HTML | French fr-FR number format with Aspose.Cells

using Aspose.Cells;
using System;
using System.Globalization;
using System.IO;

// The example loads or creates an Excel workbook, assigns a French (fr-FR) CultureInfo to the workbook settings, configures HtmlSaveOptions to include grid lines, and saves the workbook as an HTML file, ensuring that numeric values are rendered using the specified locale's formatting.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        try
        {
            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                var sheet = tempWb.Worksheets[0];
                sheet.Cells["A1"].PutValue(1234.56); // sample numeric value
                tempWb.Save(inputPath);
            }

            // Load the Excel workbook
            var workbook = new Workbook(inputPath);

            // Set the workbook's culture to the desired locale (e.g., French - France)
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Configure HTML save options (grid lines optional)
            var htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportGridLines = true
            };

            // Export the workbook to HTML with locale‑specific number formatting
            workbook.Save(outputPath, htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
