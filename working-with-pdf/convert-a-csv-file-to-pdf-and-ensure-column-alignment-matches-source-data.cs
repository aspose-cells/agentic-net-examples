// Title: How to convert a CSV file to PDF with exact column widths using Aspose.Cells for .NET (C#)
// Common Searches: c# aspnet convert csv to pdf keep column widths aspose.cells | aspose.cells pdfsaveoptions preserve column alignment when exporting csv | load csv with loadoptions and auto fit columns before saving as pdf in c# | how to prevent column wrapping in pdf generated from csv using aspose.cells | c# example to export csv to pdf with original column sizes using aspose.cells
// Tags: CSV to PDF conversion using Aspose.Cells for .NET | auto-fit worksheet columns prior to PDF generation | PdfSaveOptions AllColumnsInOnePagePerSheet configuration | LoadOptions specifying LoadFormat.Csv in Aspose.Cells | handling missing input CSV file in C# conversion script | preserve original column dimensions in PDF output

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a CSV file with LoadOptions, auto‑fits the worksheet columns, sets PdfSaveOptions.AllColumnsInOnePagePerSheet to false to keep original widths, and saves the workbook as a PDF while handling missing files gracefully.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.csv";
            const string outputPath = "output.pdf";

            // Verify that the input CSV file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the CSV file into a workbook using LoadOptions to specify CSV format.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Auto‑fit columns so the widths match the content from the CSV.
            Worksheet sheet = workbook.Worksheets[0];
            sheet.AutoFitColumns();

            // Configure PDF save options. AllColumnsInOnePagePerSheet = false keeps original widths.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                AllColumnsInOnePagePerSheet = false
                // Note: RenderGridLines property is not available in this version of Aspose.Cells.
            };

            // Save the workbook as a PDF file.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
