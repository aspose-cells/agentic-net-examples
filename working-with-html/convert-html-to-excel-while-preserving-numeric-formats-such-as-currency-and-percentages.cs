// Title: Convert HTML tables to an Excel workbook in C# while preserving currency and percentage formats with Aspose.Cells
// AI Prompts: Generate C# code that loads an HTML file using Aspose.Cells HtmlLoadOptions with numeric conversion enabled and saves it as an XLSX workbook. | Show how to turn HTML strings like "$1,234.56" and "12%" into true numeric cells during the HTML‑to‑Excel conversion. | Demonstrate applying custom number‑format strings after conversion so that currency and percentage cells keep their display style.
// Common Searches: Aspose.Cells C# convert HTML to XLSX keep currency formatting | How to preserve percentage values when loading HTML into a workbook with Aspose.Cells | Convert numeric strings from HTML tables to Excel numbers using HtmlLoadOptions | C# example for HTML to Excel conversion with numeric data type preservation | Load HTML file as workbook and retain numeric cell types Aspose.Cells .NET
// Tags: HTML to Excel numeric conversion Aspose.Cells | numeric conversion flag Aspose.Cells | preserve currency format Aspose.Cells | percentage values as Excel numbers | load HTML workbook with numeric cells

using System;
using System.IO;
using Aspose.Cells;

// C# program that reads an HTML file with Aspose.Cells HtmlLoadOptions (numeric conversion enabled) and saves it as an XLSX workbook, ensuring that currency and percentage strings become proper numeric cells with appropriate formatting.
class HtmlToExcelConverter
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlPath = "input.html";

        // Path for the generated Excel file
        string excelPath = "output.xlsx";

        // Verify that the input HTML file exists
        if (!File.Exists(htmlPath))
        {
            Console.WriteLine($"Error: The file \"{htmlPath}\" was not found.");
            return;
        }

        try
        {
            // Configure HTML load options to convert numeric strings to numeric cells
            HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html)
            {
                ConvertNumericData = true // Convert numeric strings to numeric cells
            };

            // Load the HTML content into a workbook using the configured options
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Save the workbook as an Excel file (XLSX)
            workbook.Save(excelPath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion successful. Excel file saved to \"{excelPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during conversion: {ex.Message}");
        }
    }
}
