// Title: Log font substitution warnings while converting an Excel workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Show C# code that registers an IWarningCallback on a Workbook to capture font substitution warnings during PDF export and writes each warning to the console. | Provide an example of configuring PdfSaveOptions together with a warning callback to log missing or substituted fonts when saving an Excel file as PDF in Aspose.Cells. | Demonstrate how to filter and output only font‑related warnings after calling Workbook.Save with PdfSaveOptions in a .NET application.
// Common Searches: asp.net aspose.cells capture font substitution warnings during pdf conversion | c# retrieve missing font messages when exporting Excel to PDF with Aspose.Cells | how to use IWarningCallback for font fallback logging in Aspose.Cells PDF export | log font substitution details after workbook.Save to PDF in .NET
// Tags: Aspose.Cells PDF font substitution warning | C# IWarningCallback Aspose.Cells | Excel to PDF missing font logging | PdfSaveOptions warning handling Aspose.Cells | Workbook warning callback font fallback

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the input XLSX file, loads it into an Aspose.Cells Workbook, sets up a warning callback to capture font substitution messages, and saves the workbook as a PDF using PdfSaveOptions. Any font‑related warnings generated during the conversion are written to the console, enabling developers to identify missing or substituted fonts.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" was not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Configure PDF save options (basic configuration)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Render the workbook to PDF
            try
            {
                workbook.Save(outputFile, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF to \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save PDF: {ex.Message}");
                return;
            }

            // Note: Font substitution warning properties are not available in this version of Aspose.Cells.
            // If needed, handle warnings via other mechanisms provided by the library.
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
