// Title: Convert a CSV file to PDF with Aspose.Cells in C# and understand why Office Add‑In form fields are not preserved
// Common Searches: how to keep Office Add‑In form fields when converting CSV to PDF using Aspose.Cells C# | Aspose.Cells PDF conversion from CSV preserving interactive controls | which Aspose.Cells version adds support for PreserveFormFields property | C# example for loading a CSV workbook and saving as PDF with PdfSaveOptions | why does PreserveFormFields not work when saving a CSV workbook as PDF in Aspose.Cells
// Tags: Aspose.Cells CSV to PDF conversion | PdfSaveOptions PreserveFormFields limitation | Office Add‑In form field preservation Aspose.Cells | C# workbook load CSV Aspose.Cells | upgrade Aspose.Cells for form field support

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The sample loads a CSV file into an Aspose.Cells Workbook, configures PdfSaveOptions, and saves the workbook as a PDF. It notes that the PreserveFormFields property is unavailable in the current Aspose.Cells version, advising an upgrade if form‑field preservation from Office Add‑Ins is required.
class Program
{
    static void Main()
    {
        const string inputPath = "input.csv";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input CSV file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the CSV file into a workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // The PreserveFormFields property is not available in the current Aspose.Cells version.
            // If preserving form fields is required, upgrade to a version that supports it.

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
