// Title: Check and set the Title built‑in document property before exporting an Aspose.Cells workbook to PDF (C#)
// Common Searches: how to set a default Title property in Aspose.Cells before saving to PDF in C# | Aspose.Cells C# verify built‑in document properties are not empty prior to export | prevent empty Title built‑in property when converting Excel to PDF with Aspose.Cells | C# code to check and assign workbook Title property using Aspose.Cells | Aspose.Cells save workbook as PDF after ensuring required document properties
// Tags: check built-in document properties Aspose.Cells C# | set default workbook title Aspose.Cells PDF export | validate Title property before SaveFormat.Pdf | Aspose.Cells document property handling C# | export Excel to PDF with property validation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads or creates a workbook, ensures the built‑in Title property has a non‑empty value (assigning a default if needed), and then saves the workbook as a PDF using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample data");
            }

            // Ensure the Title built‑in property is set.
            var titleProp = workbook.BuiltInDocumentProperties["Title"];
            if (titleProp != null && string.IsNullOrWhiteSpace(titleProp.Value?.ToString()))
            {
                titleProp.Value = "Untitled Workbook";
            }

            // Export to PDF.
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
