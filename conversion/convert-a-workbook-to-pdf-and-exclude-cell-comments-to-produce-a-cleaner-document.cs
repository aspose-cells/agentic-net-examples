// Title: Convert an Excel workbook to PDF without rendering cell comments using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells and saves it as a PDF while ensuring cell comments are not included in the output. | Show how to configure PdfSaveOptions in Aspose.Cells to suppress comments during Excel‑to‑PDF conversion. | Create a robust C# example that checks for the source workbook file, applies PdfSaveOptions, and handles exceptions when exporting to PDF without comments. | Demonstrate how to verify that the generated PDF excludes Excel comment annotations using Aspose.Cells.
// Common Searches: aspnet convert excel file to pdf without comments using aspose.cells | c# export workbook to pdf hide cell comments | pdfsaveoptions commentvisibility false asp.net | how to prevent comment boxes from appearing in pdf generated from excel with aspose | skip annotations when saving xlsx as pdf in c#
// Tags: Aspose.Cells PDF export exclude comments | C# PdfSaveOptions hide cell annotations | Excel to PDF conversion without comment rendering | file existence validation Aspose.Cells | exception handling workbook to pdf conversion

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that the input .xlsx file exists, loads it with Aspose.Cells, uses PdfSaveOptions (which by default omit cell comments) to save the workbook as a PDF, and includes error handling for a reliable Excel‑to‑PDF conversion without comment annotations.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default settings do not render comments)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file with the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
