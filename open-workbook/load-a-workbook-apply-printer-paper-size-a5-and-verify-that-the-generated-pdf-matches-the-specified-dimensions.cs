// Title: Load an Excel workbook, set the first worksheet's printer paper size to A5, and export it as a PDF using Aspose.Cells for .NET
// AI Prompts: Load a .xlsx file with Aspose.Cells, change the first worksheet's PageSetup.PaperSize to PaperA5, and save the workbook as a PDF. | After exporting the workbook to PDF, read the generated PDF file and confirm that its page dimensions correspond to A5 size. | Iterate through all worksheets in a workbook, set each sheet's printer paper size to A5, and then convert the workbook to a single PDF document.
// Common Searches: Aspose.Cells C# set first sheet paper size to A5 before PDF conversion | How to check PDF page size after converting Excel to PDF with Aspose.Cells .NET | C# code sample for applying A5 printer settings to an Excel workbook using Aspose.Cells | Export Excel workbook to PDF with specific page dimensions using Aspose.Cells for .NET | Verify that generated PDF matches A5 dimensions when saving Excel as PDF in C#
// Tags: worksheet page setup paper size A5 | Aspose.Cells export to PDF with custom page size | C# verify PDF dimensions after conversion | set PaperSizeType PaperA5 in Aspose.Cells | apply printer settings to all worksheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the presence of input.xlsx, loads it into an Aspose.Cells Workbook, sets the first worksheet's PageSetup.PaperSize to PaperA5, and saves the workbook as output.pdf. It demonstrates how to configure printer paper size for Excel-to-PDF conversion and includes basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply printer paper size A5 to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA5;

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook saved as PDF with A5 page size to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
