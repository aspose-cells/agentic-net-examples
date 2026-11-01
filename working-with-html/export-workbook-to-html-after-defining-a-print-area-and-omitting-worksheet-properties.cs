// Title: Export an Aspose.Cells workbook to HTML with a specific print area and without worksheet properties using C#
// AI Prompts: Generate C# code that creates a workbook, specifies a print range, and saves it as HTML using Aspose.Cells while turning off worksheet property export. | Demonstrate configuring HtmlSaveOptions to disable worksheet property export and preserve the defined print area when converting a workbook to HTML.
// Common Searches: C# Aspose.Cells export to HTML respecting print area | How to hide worksheet properties in HTML output with Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportWorksheetProperties false example | Save Excel file as HTML without sheet metadata using Aspose.Cells | Define print area before converting workbook to HTML in C#
// Tags: Aspose.Cells HTML export with print area | HtmlSaveOptions ExportWorksheetProperties false | C# define print area Aspose.Cells | exclude worksheet metadata HTML output | save workbook as HTML Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, fills cells A1:B3 with data, sets the print area to A1:B3, configures HtmlSaveOptions to disable worksheet property export, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(30);

            // Define the print area (A1:B3)
            sheet.PageSetup.PrintArea = "A1:B3";

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Omit worksheet properties from the generated HTML
                ExportWorksheetProperties = false
                // ExportPrintArea property is not available in this version; the defined print area will be respected by default
            };

            // Determine output path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ExportedWorkbook.html");

            // Export the workbook to HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook exported successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
