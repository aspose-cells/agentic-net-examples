// Title: Save an Excel workbook as PDF with Aspose.Cells and enforce PDF version 1.4 for legacy viewers (C#)
// AI Prompts: Generate C# code that loads an .xlsx file using Aspose.Cells, configures PdfSaveOptions to produce a PDF version 1.4, and saves the workbook as a PDF. | Show how to apply reflection in C# to assign the PdfVersion property of PdfSaveOptions to the enum value Pdf14 only when the property exists. | Write error‑handling logic that continues the conversion if setting PdfVersion fails, falling back to the default PDF version.
// Common Searches: asp.net aspose.cells export excel to pdf with specific pdf version 1.4 | c# set pdf compatibility level when converting xlsx to pdf using aspose.cells | use reflection to set PdfVersion enum in PdfSaveOptions aspose.cells | pdf version 1.4 output for older viewers with aspose.cells pdfsaveoptions | fallback when PdfVersion property is missing in Aspose.Cells PdfSaveOptions
// Tags: Aspose.Cells PdfSaveOptions set PDF version | C# export Excel to PDF with specific compatibility | reflection assign PdfVersion enum Aspose.Cells | PDF 1.4 legacy viewer compatibility Aspose | handle missing PdfVersion property in PdfSaveOptions

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // for PdfSaveOptions

namespace AsposeCellsExample
{
    // Loads an Excel file with Aspose.Cells, uses reflection to set PdfSaveOptions.PdfVersion to Pdf14 when available, and saves the workbook as a PDF compatible with older PDF viewers.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Attempt to set PDF version to 1.4 if the property exists
                var pdfVersionProp = typeof(PdfSaveOptions).GetProperty("PdfVersion");
                if (pdfVersionProp != null && pdfVersionProp.CanWrite)
                {
                    try
                    {
                        // Parse the enum value by name without directly referencing the enum type
                        var enumValue = Enum.Parse(pdfVersionProp.PropertyType, "Pdf14");
                        pdfVersionProp.SetValue(pdfOptions, enumValue);
                    }
                    catch
                    {
                        // If parsing fails (e.g., enum value not present), ignore and continue
                    }
                }

                // Save the workbook as PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
