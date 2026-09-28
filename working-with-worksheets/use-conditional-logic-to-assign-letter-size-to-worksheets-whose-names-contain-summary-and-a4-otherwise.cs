// Title: Assign Letter paper size to worksheets containing "Summary" and A4 to all other sheets using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that opens an Excel workbook with Aspose.Cells, iterates over each worksheet, and sets Worksheet.PageSetup.PaperSize to PaperLetter when the sheet name includes the word "Summary" (case‑insensitive), otherwise sets it to PaperA4, then saves the file. | Generate C# code that conditionally changes the page setup of every worksheet in a workbook: apply Letter size to sheets whose names contain "Summary" and A4 size to the remaining sheets, using Aspose.Cells' PageSetup API.
// Common Searches: Aspose.Cells set paper size based on worksheet name C# | How to apply Letter page size to summary tabs in an Excel file using .NET | Conditional page setup for multiple worksheets with Aspose.Cells | C# change worksheet PageSetup.PaperSize to A4 for non‑summary sheets
// Tags: conditional worksheet paper size Aspose.Cells | PageSetup.PaperSize assignment by worksheet name | Letter size for summary tabs using Aspose.Cells | A4 size for non‑summary sheets Aspose.Cells | modify workbook page setup and save with Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program loads an existing workbook, checks each worksheet name, assigns PaperLetter size to sheets containing "Summary" (case‑insensitive) and PaperA4 to all other sheets, ensures the output directory exists, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and set paper size based on name
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name.IndexOf("Summary", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Assign Letter paper size
                    sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
                }
                else
                {
                    // Assign A4 paper size
                    sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
                }
            }

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
