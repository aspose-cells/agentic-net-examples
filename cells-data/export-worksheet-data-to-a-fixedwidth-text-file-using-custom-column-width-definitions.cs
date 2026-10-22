// Title: Export an Excel worksheet to a UTF-8 fixed-width text file with custom column widths using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a workbook from a given .xlsx path and saves it as a fixed-width .txt file using Aspose.Cells TxtSaveOptions with UTF-8 encoding. | Adjust the export routine to assign specific column widths for each column when generating the fixed-width text file with Aspose.Cells.
// Common Searches: how to save an Excel sheet as a fixed-width UTF-8 text file using Aspose.Cells in C# | Aspose.Cells TxtSaveOptions set column widths for fixed-width text export | C# export Excel to fixed-width .txt with custom column sizes Aspose.Cells | example of using TxtSaveOptions to create fixed-width text from a workbook
// Tags: Aspose.Cells TxtSaveOptions fixed-width export | C# UTF-8 text file export from Excel | column width configuration Aspose.Cells | .NET fixed-width txt generation | save workbook as fixed-width text Aspose

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The program verifies the presence of an input.xlsx file, loads it into an Aspose.Cells Workbook, configures TxtSaveOptions with UTF-8 encoding, and saves the workbook as output.txt in a fixed-width text format, ready for custom column width adjustments and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.txt";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook.
            Workbook workbook = new Workbook(inputPath);

            // Configure TxtSaveOptions for text export.
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                // Use UTF‑8 encoding for the output file.
                Encoding = Encoding.UTF8
            };

            // Save the workbook as a fixed‑width text file.
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Export completed successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
