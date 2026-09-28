// Title: Split an Excel workbook into individual PDF files per worksheet with indexed and file‑system‑safe names using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells, iterates through all worksheets, creates a temporary workbook for each sheet, and saves it as a PDF named "{index}_{sanitizedSheetName}.pdf" in a target folder. | Write a C# helper method that converts an invalid worksheet name into a file‑system‑compatible string and use it to build custom PDF filenames when exporting worksheets with Aspose.Cells.
// Common Searches: Aspose.Cells C# export each worksheet to a separate PDF with custom naming | how to save Excel sheets as individual PDFs using Aspose.Cells .NET | C# generate PDF per sheet with index prefix and safe filename | remove invalid characters from worksheet name for PDF export Aspose.Cells | split workbook into multiple PDFs programmatically in .NET
// Tags: worksheet to PDF export Aspose.Cells | indexed PDF filenames for Excel sheets | file‑system safe worksheet names C# | single-sheet workbook export Aspose.Cells | batch export Excel worksheets to PDFs .NET

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, iterates through its worksheets, copies each sheet into a temporary workbook, sanitizes the sheet name, and saves each as a PDF file named with a two‑digit index and a file‑system‑safe sheet name in the specified output directory.
class SplitWorkbookToPdf
{
    static void Main()
    {
        try
        {
            // Path to the source Excel workbook
            string sourceFile = @"C:\Input\Workbook.xlsx";

            // Verify that the source file exists
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Source file not found: {sourceFile}");
                return;
            }

            // Folder where individual PDF files will be saved
            string outputFolder = @"C:\Output\PdfSheets";

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Load the workbook
            Workbook workbook = new Workbook(sourceFile);

            // Iterate through each worksheet in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                // Create a temporary workbook containing only the current worksheet
                Workbook tempWorkbook = new Workbook();

                // The new workbook contains a default empty sheet; copy the source sheet into it
                Worksheet tempSheet = tempWorkbook.Worksheets[0];
                tempSheet.Copy(sheet);
                tempSheet.Name = sheet.Name; // preserve original sheet name

                // Build a custom file name: "SheetIndex_SheetName.pdf"
                string safeSheetName = MakeFileNameSafe(sheet.Name);
                string pdfFileName = $"{i + 1:D2}_{safeSheetName}.pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Save the temporary workbook as PDF
                tempWorkbook.Save(pdfPath, SaveFormat.Pdf);
            }

            Console.WriteLine("Workbook split and saved as PDFs successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to remove invalid filename characters
    private static string MakeFileNameSafe(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(c, '_');
        }
        return name;
    }
}
