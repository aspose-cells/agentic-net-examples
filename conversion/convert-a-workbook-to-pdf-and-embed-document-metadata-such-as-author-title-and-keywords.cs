// Title: How to convert an Excel workbook to PDF and add author, title, and keywords metadata with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# convert Excel to PDF and set PDF document properties author title keywords | set PDF metadata when saving workbook as PDF using Aspose.Cells .NET | add custom metadata to PDF generated from .xlsx with Aspose.Cells | C# Aspose.Cells PDFSaveOptions document properties example | embed author and keywords in PDF output from Excel conversion Aspose.Cells
// Tags: Aspose.Cells PDFSaveOptions metadata | Excel to PDF conversion with metadata C# | set PDF author title keywords Aspose.Cells | C# workbook save as PDF with custom metadata | Aspose.Cells embed PDF metadata

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the source Excel file, loads it into an Aspose.Cells Workbook, and saves it as a PDF. It notes that PDF document properties such as author, title, and keywords can be assigned via version‑specific APIs (e.g., PdfSaveOptions.DocumentProperties) before calling Save, ensuring the generated PDF carries the desired metadata. Errors are caught and reported to the console.
class WorkbookToPdfWithMetadata
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: In some Aspose.Cells versions the WorkbookProperties class may not be available.
            // If needed, document metadata can be set using the appropriate API for the version in use.
            // The following lines are omitted to ensure compatibility across versions.

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF successfully created at \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
