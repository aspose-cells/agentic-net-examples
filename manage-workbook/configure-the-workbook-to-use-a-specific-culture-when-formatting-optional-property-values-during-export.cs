// Title: Setting workbook CultureInfo to French (France) for culture‑aware date formatting during PDF export with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set workbook culture to French before PDF conversion | export PDF from Aspose.Cells with French date format in .NET | configure Settings.CultureInfo for optional property values in Aspose.Cells | apply locale‑specific date formatting when converting Excel to PDF using Aspose.Cells | how to use CultureInfo with Aspose.Cells PDFSaveOptions
// Tags: Aspose.Cells workbook Settings.CultureInfo | PDF export culture-specific formatting | French locale date style Aspose.Cells | culture-aware Excel to PDF conversion .NET | custom date format dd MMMM yyyy Aspose.Cells

using Aspose.Cells;
using System;
using System.Globalization;
using System.IO;

// The example creates a Workbook, assigns Settings.CultureInfo to "fr-FR", writes the current date to a cell, applies a custom date style (dd MMMM yyyy) that follows the French locale, and saves the workbook as a PDF using PdfSaveOptions, demonstrating culture‑aware formatting during export.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the workbook's culture to French (France)
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Add sample data to demonstrate culture‑sensitive formatting
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(DateTime.Now); // Insert current date/time

            // Apply a date format that will be rendered according to the culture
            Style style = cell.GetStyle();
            style.Custom = "dd MMMM yyyy";
            cell.SetStyle(style);

            // Configure PDF export options (no culture property needed)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure the output directory exists
            string outputPath = "ExportedWorkbook.pdf";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the workbook to PDF using the specified culture settings
            workbook.Save(outputPath, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
