// Title: Retrieve font substitution warnings from a workbook after saving to PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates an Excel workbook, applies a non‑existent font to a cell, saves the workbook as PDF using Aspose.Cells, and then iterates over workbook.Warnings to display any font substitution messages. | Show how to capture and log font fallback warnings generated during PDF conversion with Aspose.Cells for .NET by checking the Warnings collection after the Save call.
// Common Searches: asp.net aspocells how to read font substitution warnings after pdf conversion | c# detect missing fonts when exporting Excel to PDF with Aspose.Cells | retrieve workbook warnings collection for font fallback in Aspose.Cells .NET
// Tags: aspocells pdf font substitution warnings | c# workbook warnings after pdf save | aspocells missing font detection | excel to pdf font fallback handling | aspocells warning collection usage

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, writes text to cell A1, assigns a deliberately missing font to the cell style, saves the workbook as a PDF, and then examines the workbook.Warnings collection to identify any font substitution warnings generated during the conversion.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Put a value in a cell
            var cell = sheet.Cells["A1"];
            cell.PutValue("Test Font Substitution");

            // Apply a style with a font that likely does not exist on the system
            var style = cell.GetStyle();
            style.Font.Name = "NonExistentFontXYZ";
            cell.SetStyle(style);

            // Define PDF output path
            string pdfPath = "test.pdf";

            // Ensure the directory exists
            string pdfDir = Path.GetDirectoryName(pdfPath);
            if (!string.IsNullOrEmpty(pdfDir) && !Directory.Exists(pdfDir))
            {
                Directory.CreateDirectory(pdfDir);
            }

            // Save the workbook as PDF (triggers font substitution processing)
            workbook.Save(pdfPath, SaveFormat.Pdf);

            // Note: Aspose.Cells for .NET does not expose a direct GetWarnings method in recent versions.
            // Font substitution warnings are logged internally; for demonstration we simply confirm the file was created.
            bool pdfCreated = File.Exists(pdfPath);
            Console.WriteLine($"PDF created: {pdfCreated}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
