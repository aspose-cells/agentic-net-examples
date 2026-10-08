// Title: Export Excel to HTML with German locale number formatting using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, sets Workbook.Settings.CultureInfo to a specified locale (e.g., de-DE), applies a custom numeric format to a defined cell range, and saves the workbook as HTML with images embedded as Base64 using Aspose.Cells. | Adapt the example to export only the active worksheet to HTML while keeping the workbook's culture‑specific number separators and embedding images as Base64. | Create a reusable C# method ExportToHtml(inputPath, outputPath, cultureCode, numberFormat) that applies the given culture and numeric format to the workbook and writes an HTML file using Aspose.Cells HtmlSaveOptions.
// Common Searches: how to export Excel to HTML with German number formatting using Aspose.Cells .NET | Aspose.Cells set workbook culture info before HTML conversion C# | apply custom numeric format to a range and export to HTML with embedded images Aspose.Cells | export entire workbook to HTML preserving locale specific separators Aspose.Cells | C# Aspose.Cells HtmlSaveOptions culture-specific formatting example
// Tags: Aspose.Cells HTML export with culture-specific formatting | C# set workbook CultureInfo Aspose.Cells | custom numeric format range Aspose.Cells | embed images Base64 HTML Aspose.Cells | German locale number separators Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Loads InputWorkbook.xlsx, sets workbook culture to German (de-DE), applies custom number format "#,##0.00" to cells A1:A10, and saves the workbook as OutputWorkbook.html with images embedded as Base64 via HtmlSaveOptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook culture to the desired locale (e.g., German - Germany)
            workbook.Settings.CultureInfo = new CultureInfo("de-DE");

            // Apply a custom number format to the range A1:A10 on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Style style = workbook.CreateStyle();
            style.Custom = "#,##0.00"; // Separator follows the workbook culture
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            sheet.Cells.CreateRange("A1:A10").ApplyStyle(style, flag);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = false,   // Export the whole workbook
                ExportImagesAsBase64 = true          // Embed images as Base64 strings
                // Note: ExportCultureInfo and ExportEmbeddedCss are not available in the current API version
            };

            // Save the workbook as an HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
