// Title: Load an Excel workbook with Aspose.Cells, set all worksheets to A3 paper size, and save as PDF using C#
// Common Searches: Aspose.Cells C# set A3 paper size for every worksheet before PDF export | How to change page setup to A3 for all sheets using Aspose.Cells .NET | Convert XLSX to PDF with A3 page layout using Aspose.Cells in C# | C# code to apply printer paper size A3 to a workbook and generate PDF with Aspose.Cells
// Tags: set worksheet paper size A3 Aspose.Cells | export workbook to PDF C# | page setup configuration Aspose.Cells | apply printer settings to Excel workbook | convert XLSX to PDF with custom page size

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file into an Aspose.Cells Workbook, iterates through each worksheet to set the PageSetup.PaperSize to PaperSizeType.PaperA3, and then saves the workbook as a PDF file, handling missing input files and runtime exceptions.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string inputPath = "input.xlsx";

        // Path for the resulting PDF file
        string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook directly from the file
            Workbook workbook = new Workbook(inputPath);

            // Set the printer paper size to A3 for all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
            }

            // Save the workbook as a PDF document
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
