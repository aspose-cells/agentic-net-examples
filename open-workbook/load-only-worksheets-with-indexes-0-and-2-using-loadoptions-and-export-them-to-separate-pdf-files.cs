// Title: Export only worksheets 0 and 2 from an Excel workbook to separate PDF files using Aspose.Cells for .NET
// Common Searches: C# how to save selected Excel worksheets as individual PDF files with Aspose.Cells | Aspose.Cells export specific sheet indexes to PDF in .NET | Create temporary workbook with one sheet and convert to PDF using Aspose.Cells | Load only certain worksheets from Excel and generate separate PDFs using Aspose.Cells | Aspose.Cells PDFSaveOptions OnePagePerSheet example C#
// Tags: export worksheet by index to PDF Aspose.Cells | OnePagePerSheet PDFSaveOptions usage | add copy of worksheet to temporary workbook | validate sheet index range Aspose.Cells | C# generate individual PDF files from Excel sheets

using System;
using System.IO;
using Aspose.Cells;

// The C# program verifies the source Excel file, loads it with Aspose.Cells, selects worksheets at indexes 0 and 2, copies each into a temporary workbook, configures PdfSaveOptions with OnePagePerSheet enabled, and saves each sheet as a separate PDF (Sheet0.pdf, Sheet2.pdf). It includes index range checks and exception handling for robust execution.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the entire workbook (no custom filter needed)
            Workbook workbook = new Workbook(inputPath);

            // Define which worksheet indexes should be exported to PDF
            int[] targetSheets = new int[] { 0, 2 };

            foreach (int sheetIndex in targetSheets)
            {
                // Ensure the requested sheet index is within the workbook range
                if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
                {
                    Console.WriteLine($"Warning: Sheet index {sheetIndex} is out of range.");
                    continue;
                }

                try
                {
                    string pdfFileName = $"Sheet{sheetIndex}.pdf";

                    // Create a temporary workbook containing only the selected sheet
                    Workbook tempWb = new Workbook();
                    tempWb.Worksheets.Clear();

                    // Add a copy of the selected worksheet by name (correct overload)
                    string sheetName = workbook.Worksheets[sheetIndex].Name;
                    tempWb.Worksheets.AddCopy(sheetName);

                    // Configure PDF save options
                    PdfSaveOptions pdfOptions = new PdfSaveOptions
                    {
                        OnePagePerSheet = true
                    };

                    // Save the selected worksheet as a PDF
                    tempWb.Save(pdfFileName, pdfOptions);
                    Console.WriteLine($"Saved sheet {sheetIndex} to '{pdfFileName}'.");
                }
                catch (Exception exSheet)
                {
                    Console.WriteLine($"Failed to save sheet {sheetIndex}: {exSheet.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
