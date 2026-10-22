// Title: Convert an Excel workbook to PDF with cell comments retained as PDF annotations using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, verifies the file exists, and saves it as a PDF while automatically converting cell comments into PDF annotations. | Show how to configure PdfSaveOptions in Aspose.Cells to customize the rendering of comment annotations during Excel‑to‑PDF conversion. | Create robust error‑handling for the workbook‑to‑PDF process that logs missing input files and captures exceptions thrown by Aspose.Cells.
// Common Searches: Aspose.Cells .NET how to keep Excel cell comments when saving as PDF | C# convert .xlsx to PDF preserving comments as annotations | PdfSaveOptions comment handling Aspose.Cells example | Check if Excel file exists before converting to PDF with Aspose.Cells | Export Excel workbook to PDF with comment annotations using Aspose.Cells
// Tags: Aspose.Cells Excel to PDF conversion with comments | PdfSaveOptions comment annotation settings | C# workbook file existence validation | Export XLSX to PDF preserving annotations | Aspose.Cells error handling during PDF export

using System;
using System.IO;
using Aspose.Cells;

// // This C# program checks that 'input.xlsx' exists, loads it with Aspose.Cells, and saves it as 'output.pdf' using PdfSaveOptions, which automatically includes cell comments as PDF annotations.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (comments are preserved by default)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
