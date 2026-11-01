// Title: C# – Load an Excel workbook, set worksheet printer paper size to Letter, and export to PDF with Aspose.Cells
// Common Searches: C# Aspose.Cells set worksheet printer paper size to Letter before saving as PDF | How to export an Excel file to PDF with Letter page size using Aspose.Cells in .NET | Aspose.Cells C# retrieve PDF page count after converting a workbook to PDF | Validate input Excel file exists before loading it with Aspose.Cells in C# | Save Excel workbook as PDF with custom page setup using Aspose.Cells .NET
// Tags: Aspose.Cells set paper size Letter | Aspose.Cells workbook to PDF export | custom page setup Excel PDF Aspose.Cells | C# file existence check Aspose.Cells | PDF page count extraction after Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that 'input.xlsx' exists, loads it into an Aspose.Cells Workbook, sets the first worksheet's printer paper size to Letter, and saves the workbook as 'output.pdf'. It notes that page count must be obtained from the generated PDF, and includes basic exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Apply printer paper size Letter
            sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;

            // NOTE: The GetPagesCount method is not available in the current Aspose.Cells version.
            // If page count information is required, it can be obtained after saving to PDF
            // using a PDF library or by inspecting the generated PDF file.
            // For demonstration, we will proceed without retrieving the page count here.

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display an error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
