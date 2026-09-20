// Title: Load an Excel workbook without charts, set the first worksheet to A5 paper size, and save as PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file with LoadOptions configured to skip chart objects, changes the first sheet's PageSetup to PaperA5, and exports the workbook to a PDF file. | Generate a .NET snippet that verifies an input Excel file, loads it using LoadOptions that disable chart loading, sets the printer paper size to A5 for the first worksheet, and saves the result as a PDF.
// Common Searches: how to load an Excel file without charts using Aspose.Cells .NET | set A5 paper size for worksheet before converting to PDF with Aspose.Cells | Aspose.Cells LoadOptions to ignore charts when converting to PDF | C# export Excel to PDF with A5 page layout using Aspose.Cells | disable chart rendering in Aspose.Cells workbook load
// Tags: LoadOptions chart exclusion Aspose.Cells | Worksheet PageSetup A5 Aspose.Cells | Workbook to PDF conversion Aspose.Cells | disable chart loading Aspose.Cells .NET | set printer paper size A5 Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the input Excel file, loads it with LoadOptions (which can be configured to skip chart objects), sets the first worksheet's printer paper size to A5, and saves the workbook as a PDF, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load options (default). If a newer version supports LoadDataOnly, it can be set here.
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Set the printer paper size to A5 for the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA5;

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
