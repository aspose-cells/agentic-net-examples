// Title: Generate separate PDF files for each worksheet in an Excel workbook with custom safe filenames using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export each worksheet to its own PDF file with custom naming | How to split an Excel workbook into multiple PDFs per sheet using Aspose.Cells .NET | Save individual worksheets as PDF with safe file names in C# Aspose.Cells | Aspose.Cells create one‑page‑per‑sheet PDF for each worksheet programmatically | C# generate PDF per Excel sheet and replace invalid filename characters
// Tags: export worksheet to PDF Aspose.Cells | PdfSaveOptions OnePagePerSheet C# | create single‑sheet workbook Aspose.Cells | sanitize worksheet name for file system C# | split Excel workbook into multiple PDFs

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through its worksheets, creates a new single‑sheet workbook for each sheet, applies PdfSaveOptions with OnePagePerSheet enabled, sanitizes the worksheet name to produce a safe file name, and saves each sheet as an individual PDF named "Report_<safeSheetName>.pdf" while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source Excel workbook
            Workbook sourceWorkbook = new Workbook(inputPath);

            // Iterate through each worksheet in the source workbook
            foreach (Worksheet sheet in sourceWorkbook.Worksheets)
            {
                // Create a new workbook that will contain only the current worksheet
                Workbook singleSheetWorkbook = new Workbook();

                // Remove the default empty sheet that Aspose.Cells creates
                singleSheetWorkbook.Worksheets.Clear();

                // Add a copy of the current worksheet to the new workbook (by name)
                singleSheetWorkbook.Worksheets.AddCopy(sheet.Name);

                // Set PDF save options (e.g., one page per sheet)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true
                };

                // Build a safe file name using the worksheet name
                string safeSheetName = MakeFileNameSafe(sheet.Name);
                string outputFileName = $"Report_{safeSheetName}.pdf";

                // Save the single‑sheet workbook as a PDF file using the options
                singleSheetWorkbook.Save(outputFileName, pdfOptions);
            }
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to replace characters that cannot be used in file names
    static string MakeFileNameSafe(string name)
    {
        char[] invalidChars = Path.GetInvalidFileNameChars();
        foreach (char c in invalidChars)
        {
            name = name.Replace(c, '_');
        }
        return name;
    }
}
