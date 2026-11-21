// Title: Load an Excel workbook, set the first worksheet’s printer paper size to A5, and export it as a PDF with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set worksheet printer paper size to A5 before saving as PDF | How to export Excel to PDF with A5 page size using Aspose.Cells .NET | Verify PDF page dimensions after converting an Excel workbook with Aspose.Cells in C# | Aspose.Cells PageSetup PaperSizeType PaperA5 example code | Load workbook and apply A5 paper size for PDF output using Aspose.Cells
// Tags: Aspose.Cells worksheet A5 page setup | save workbook as PDF with custom page size | C# check PDF page size after export | PageSetup PaperSizeType enumeration usage | load Excel file and configure printer settings .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file into a Workbook, sets the first worksheet's PageSetup.PaperSize to PaperA5, saves the workbook as a PDF, confirms the A5 setting, and verifies that the PDF file was created, handling any errors that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from file
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Apply A5 paper size to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA5; // Correct enum for A5 paper size

            // Save the workbook as PDF
            try
            {
                workbook.Save(outputPath, SaveFormat.Pdf);
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save PDF: {saveEx.Message}");
                return;
            }

            // Verify that the worksheet page setup is A5
            bool isA5 = sheet.PageSetup.PaperSize == PaperSizeType.PaperA5;
            Console.WriteLine(isA5
                ? "Worksheet page size is set to A5."
                : "Worksheet page size is NOT set to A5.");

            // Confirm that the PDF file was created
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"PDF generated successfully: {outputPath}");
            }
            else
            {
                Console.WriteLine("Failed to generate PDF.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
