// Title: Export a single‑worksheet workbook to HTML with external CSS files using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells generate HTML with separate CSS folder for a single sheet | How to save a workbook as HTML and output CSS files per worksheet in .NET | Aspose.Cells HtmlSaveOptions ExportWorksheetCSSSeparately example code | Export only the active worksheet to HTML with an external stylesheet using Aspose.Cells | Create a CSS directory and export a workbook to HTML with Aspose.Cells C#
// Tags: HtmlSaveOptions ExportWorksheetCSSSeparately C# | single worksheet HTML export Aspose.Cells | external CSS folder Aspose.Cells HTML output | ExportActiveWorksheetOnly option .NET | Aspose.Cells generate separate CSS per sheet

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a one‑sheet workbook, populates it with data, ensures a CSS folder exists, configures HtmlSaveOptions to export only the active worksheet and to write a separate CSS file for the worksheet, and then saves the workbook as an HTML file.
class ExportWorkbookToHtml
{
    static void Main()
    {
        try
        {
            // Create a new workbook with a single worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(85);

            // Define the folder where the CSS file will be saved (optional)
            string cssFolder = "CssFolder";
            if (!Directory.Exists(cssFolder))
            {
                Directory.CreateDirectory(cssFolder);
            }

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportWorksheetCSSSeparately = true,   // Enable separate CSS per worksheet
                ExportActiveWorksheetOnly = true      // Export only the single sheet
                // Additional options can be set here if needed
            };

            // Export the workbook to HTML
            string htmlFile = "Workbook.html";
            workbook.Save(htmlFile, saveOptions);

            Console.WriteLine($"Workbook exported to '{htmlFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
