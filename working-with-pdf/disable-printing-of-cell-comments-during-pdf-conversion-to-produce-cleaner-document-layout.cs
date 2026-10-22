// Title: How to exclude cell comments when converting an Excel workbook to PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file using Aspose.Cells and saves it as a PDF while explicitly disabling comment rendering via PdfSaveOptions. | Show how to configure PdfSaveOptions (e.g., setting DisableComments) to produce a clean PDF layout without cell comments in a .NET application.
// Common Searches: asp.net aspose.cells export excel to pdf without cell comments | c# hide comments in pdf generated from workbook using aspose.cells | pdfsaveoptions disable comments asp.net core | convert xlsx to pdf clean layout aspose.cells c# | remove cell comment annotations from pdf output asp.net
// Tags: Aspose.Cells PDF export hide comments | C# PdfSaveOptions DisableComments | Excel to PDF conversion without cell comments | Aspose.Cells clean PDF layout | C# workbook to PDF without annotations | Aspose.Cells PDFSaveOptions comment exclusion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The sample verifies the input XLSX file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (comments are omitted by default or can be disabled via DisableComments), and saves the workbook as a PDF, resulting in a document that does not display cell comment markers.
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

            // Configure PDF save options (comments are excluded by default)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
