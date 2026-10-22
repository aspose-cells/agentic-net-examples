// Title: Convert HTML to Excel with French locale (fr-FR) for correct decimal separators using Aspose.Cells in C#
// AI Prompts: Generate C# code that loads an HTML file into an Aspose.Cells Workbook, sets the workbook's CultureInfo to French (France), and saves it as an XLSX file. | Provide a C# example that checks for the HTML source file, creates the destination folder if missing, applies French locale for number formatting, and handles exceptions during the HTML‑to‑Excel conversion with Aspose.Cells.
// Common Searches: aspocells c# convert html to xlsx with French number format | how to set workbook culture to fr-FR after loading html in Aspose.Cells | apply French locale decimal separator when exporting html tables to Excel using Aspose.Cells | c# Aspose.Cells HTML to Excel conversion respecting locale settings | error handling for html to xlsx conversion with culture info in Aspose.Cells
// Tags: HTML to XLSX conversion Aspose.Cells C# | set workbook CultureInfo fr-FR Aspose.Cells | locale‑specific decimal formatting Excel Aspose.Cells | validate input HTML file existence C# | create output directory for Aspose.Cells conversion

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The program verifies the HTML input file, ensures the output folder exists, loads the HTML into an Aspose.Cells Workbook, sets the workbook's CultureInfo to French (France) so decimal separators follow the French locale, and saves the result as an XLSX file while handling any errors.
class HtmlToExcelConverter
{
    static void Main()
    {
        // Paths for input HTML and output Excel files
        string htmlPath = @"C:\Input\sample.html";
        string excelPath = @"C:\Output\sample.xlsx";

        try
        {
            // Verify that the source HTML file exists
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException("HTML input file not found.", htmlPath);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(excelPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Load the HTML content into a workbook using the appropriate constructor
            Workbook workbook = new Workbook(htmlPath);

            // Set culture to French (France) for locale‑specific formatting
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Save the workbook in XLSX format
            workbook.Save(excelPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
