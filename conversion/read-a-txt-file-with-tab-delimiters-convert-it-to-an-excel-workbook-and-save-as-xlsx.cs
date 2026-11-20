// Title: How to convert a tab‑delimited TXT file to an XLSX workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load tab delimited text file and save as xlsx | using TxtLoadOptions to set tab separator in Aspose.Cells | convert .txt with tabs to Excel workbook programmatically in .NET | Aspose.Cells example for reading tab‑separated values and exporting to XLSX | handle missing input file when converting txt to xlsx with Aspose.Cells
// Tags: Aspose.Cells TxtLoadOptions tab separator | load tab‑delimited text into Workbook C# | save Workbook as Xlsx using Aspose.Cells | convert txt to xlsx programmatically .NET | file existence validation before Aspose.Cells conversion

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies that the input tab‑delimited text file exists, configures TxtLoadOptions with a tab character as the column separator, loads the file into an Aspose.Cells Workbook, and then saves the workbook as an XLSX file, handling any runtime exceptions.
class Program
{
    static void Main()
    {
        // Paths for input and output files
        string inputPath = "input.txt";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Configure load options for a tab‑delimited text file
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                Separator = '\t' // Set tab as the column separator
            };

            // Load the TXT file into a workbook using the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Save the workbook as an XLSX file
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
