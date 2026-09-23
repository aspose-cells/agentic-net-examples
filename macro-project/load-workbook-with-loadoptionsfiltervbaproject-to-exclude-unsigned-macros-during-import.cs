// Title: Convert an XLSM file to XLSX and strip VBA macros with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that verifies an .xlsm file exists, loads it with Aspose.Cells, and saves it as .xlsx to remove all VBA macros. | Show how to use Aspose.Cells Workbook.Save to export a macro‑enabled Excel workbook to a macro‑free XLSX format with proper exception handling. | Provide an example that loads an XLSM workbook, checks for errors, and converts it to XLSX while ensuring macros are discarded.
// Common Searches: Aspose.Cells C# convert macro enabled xlsm to xlsx without preserving VBA | How to remove VBA macros when saving an Excel workbook as XLSX using Aspose.Cells | C# code to load an .xlsm file and export to .xlsx stripping all macros Aspose.Cells | Save workbook as XLSX to discard macros Aspose.Cells .NET example
// Tags: convert xlsm to xlsx Aspose.Cells C# | remove VBA macros Aspose.Cells Workbook.Save | load macro enabled workbook Aspose.Cells | error handling file existence Aspose.Cells | macro‑free Excel export Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// The sample checks that the input XLSM file exists, loads it with Aspose.Cells Workbook, and saves it as an XLSX file, which automatically removes any VBA macros, while handling possible exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsm";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook; saving as Xlsx will automatically discard VBA macros
            Workbook workbook = new Workbook(inputPath);

            // Place for any additional workbook processing logic

            // Save the workbook without macros
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions (e.g., loading/saving errors)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
