// Title: Open an Excel 95 (.xls) workbook with Aspose.Cells for .NET by setting LoadOptions.Version = Excel95 before creating the Workbook and then save as XLSX
// AI Prompts: Write C# code that creates a LoadOptions object, sets its Version property to Excel95, and passes it to the Workbook constructor to open a legacy .xls file with Aspose.Cells. | Show a complete C# example that opens an Excel 95 file using LoadOptions, modifies a cell, and saves the workbook as an .xlsx file. | Provide error‑handling logic that verifies the source .xls exists, creates the destination folder if missing, and converts the Excel 95 workbook to XLSX using LoadOptions.Version.
// Common Searches: Aspose.Cells C# load legacy Excel 95 file with LoadOptions.Version | How to specify Excel95 version when opening .xls in Aspose.Cells | Convert old Excel 95 .xls to .xlsx using LoadOptions in .NET | C# example for opening Excel 95 workbook with Aspose.Cells LoadOptions | Set workbook version to Excel95 before saving as XLSX in Aspose.Cells
// Tags: Aspose.Cells LoadOptions Excel95 | C# open legacy XLS with Aspose.Cells | convert Excel 95 to XLSX Aspose | specify workbook version Aspose.Cells | Aspose.Cells save as XLSX

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that the input Excel 95 file exists, creates a LoadOptions instance with Version set to Excel95, opens the workbook using that LoadOptions, ensures the output directory is present, and saves the workbook as an XLSX file, reporting any errors that occur.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "Input95.xls";
            string outputPath = "Output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (format is auto‑detected)
            Workbook workbook = new Workbook(inputPath);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook in XLSX format
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
