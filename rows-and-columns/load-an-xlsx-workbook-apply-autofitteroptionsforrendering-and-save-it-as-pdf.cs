// Title: Auto‑fit all rows and columns for rendering and convert an XLSX workbook to PDF with Aspose.Cells in C#
// AI Prompts: Write C# code that loads an .xlsx file, sets AutoFitterOptions.ForRendering = true, auto‑fits rows and columns on every worksheet, and saves the workbook as a PDF using Aspose.Cells. | Create a .NET console program that checks for the existence of an input Excel file, applies AutoFitColumns and AutoFitRows with rendering options to each sheet, and outputs a PDF file. | Implement a method that receives a stream of an Excel workbook, applies AutoFitterOptions for rendering to auto‑size all rows and columns, and returns the generated PDF as a byte array.
// Common Searches: how to use AutoFitterOptions.ForRendering with Aspose.Cells to export Excel to PDF in C# | C# Aspose.Cells auto‑size rows and columns before PDF conversion | convert XLSX to PDF while preserving column widths using Aspose.Cells | apply rendering auto‑fit to all worksheets in Aspose.Cells .NET example
// Tags: auto‑fit rows columns rendering Aspose.Cells | XLSX to PDF conversion Aspose.Cells C# | AutoFitterOptions ForRendering usage | AutoFitColumns AutoFitRows PDF export | Aspose.Cells workbook rendering options

using System;
using System.IO;
using Aspose.Cells;

// Loads an XLSX workbook, applies AutoFitterOptions with ForRendering=true to auto‑fit rows and columns on each worksheet, and saves the result as a PDF file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the XLSX workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Prepare AutoFitterOptions for rendering (set ForRendering = true)
            AutoFitterOptions renderOptions = new AutoFitterOptions
            {
                ForRendering = true
            };

            // Apply AutoFitterOptions to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Auto‑fit all columns using rendering options
                sheet.AutoFitColumns(renderOptions);
                // Auto‑fit all rows using rendering options
                sheet.AutoFitRows(renderOptions);
            }

            // Save the workbook as a PDF document
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
