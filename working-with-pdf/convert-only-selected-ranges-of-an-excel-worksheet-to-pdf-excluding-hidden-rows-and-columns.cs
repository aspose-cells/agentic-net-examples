// Title: Convert a specific Excel range to PDF while ignoring hidden rows and columns using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook, defines a print area (e.g., A1:D20) on a worksheet, and saves only the visible cells of that range as a PDF with Aspose.Cells, excluding hidden rows and columns. | Show how to configure Aspose.Cells PdfSaveOptions to export a selected worksheet range to PDF while automatically skipping any hidden rows or columns in the source file.
// Common Searches: Aspose.Cells .NET export only visible cells from a defined range to PDF | C# how to skip hidden rows and columns when converting Excel to PDF with Aspose | Set print area for PDF output in Aspose.Cells and ignore hidden rows | Convert Excel range A1:D20 to PDF without hidden columns using Aspose.Cells | Save selected worksheet area as PDF while excluding hidden rows in C#
// Tags: Aspose.Cells selected range PDF export | PdfSaveOptions hide rows columns exclusion | worksheet print area PDF Aspose | C# Excel to PDF visible cells only | skip hidden rows columns Aspose.Cells conversion

using Aspose.Cells;
using System;
using System.IO;

// // Loads 'input.xlsx', sets the print area to A1:D20 on the first worksheet, and saves the visible portion of that range as 'output.pdf' using Aspose.Cells PdfSaveOptions, which automatically omits hidden rows and columns.
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

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range that should be converted to PDF
            string exportRange = "A1:D20";

            // Set the print area so only this range is exported
            sheet.PageSetup.PrintArea = exportRange;

            // Configure PDF save options (default behavior excludes hidden rows/columns in recent versions)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file using the appropriate overload
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF successfully created at '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
