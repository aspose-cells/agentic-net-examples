// Title: Export an Excel workbook to HTML with locale‑aware date formatting using Aspose.Cells for .NET
// Common Searches: how to set workbook culture before saving as HTML with Aspose.Cells in C# | Aspose.Cells export to HTML with French date format | C# generate HTML from Excel respecting regional settings using Aspose.Cells | apply built‑in date number format when converting Excel to HTML with Aspose.Cells | save entire workbook as HTML while preserving locale settings in .NET
// Tags: set workbook CultureInfo for HTML export | locale aware date formatting Aspose.Cells | HtmlSaveOptions export entire workbook | built‑in date number format 14 Aspose.Cells | C# Aspose.Cells generate HTML with regional settings | apply workbook culture to HTML conversion

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file, assigns a specific CultureInfo (e.g., fr‑FR) to the workbook, writes a DateTime value to a cell, applies the built‑in date format (number 14), and saves the workbook as an HTML file with HtmlSaveOptions, ensuring that dates in the generated HTML are formatted according to the workbook’s locale.
class Program
{
    static void Main()
    {
        // Paths to the input Excel file and the output HTML file.
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input file exists.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // ------------------------------------------------------------
            // 1. Set the workbook's locale (culture) that will be used for
            //    formatting dates, numbers, etc.
            // ------------------------------------------------------------
            // Example: French (France) locale. Change to any valid culture name
            //          such as "en-US", "de-DE", "ja-JP", etc.
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // ------------------------------------------------------------
            // 2. Insert a sample date value to demonstrate locale‑aware formatting.
            // ------------------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put a DateTime value into cell A1.
            cells["A1"].PutValue(DateTime.Now);

            // Apply a built‑in date number format (14 = "m/d/yyyy" in invariant culture,
            // but the actual displayed string will respect the workbook's CultureInfo).
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // Built‑in date format.
            cells["A1"].SetStyle(dateStyle);

            // ------------------------------------------------------------
            // 3. Save the workbook as HTML while respecting the locale.
            // ------------------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export the whole workbook (not only the active worksheet).
                ExportActiveWorksheetOnly = false
                // Additional options can be set here if needed.
            };

            // Save the workbook to an HTML file.
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file generated successfully at: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred during processing:");
            Console.WriteLine(ex.Message);
        }
    }
}
