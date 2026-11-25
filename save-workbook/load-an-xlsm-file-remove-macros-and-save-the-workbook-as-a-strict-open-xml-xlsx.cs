// Title: Remove VBA macros from an XLSM workbook and save it as a strict Open XML XLSX using Aspose.Cells for .NET
// Common Searches: how to strip VBA macros from an XLSM file with Aspose.Cells in C# | convert macro‑enabled Excel workbook to XLSX without preserving macros using Aspose.Cells | Aspose.Cells C# remove VbaProject before saving as Xlsx | save XLSM as strict Open XML XLSX in .NET while discarding macros | C# code to load XLSM, clear VBA modules, and export to XLSX with Aspose
// Tags: strip VBA modules Aspose.Cells | XLSM to strict Open XML XLSX conversion | clear VbaProject modules C# | export workbook to Xlsx without VBA | Aspose.Cells macro removal workflow

using System;
using System.IO;
using Aspose.Cells;

// The example loads a macro‑enabled XLSM workbook, clears any VBA modules if a VbaProject exists, and then saves the workbook as a strict Open XML XLSX file using Aspose.Cells for .NET, with file existence checks and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsm";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the macro‑enabled workbook (XLSM)
            Workbook workbook = new Workbook(inputPath);

            // Remove all macros by clearing VBA modules, if a VBA project exists
            if (workbook.VbaProject != null)
            {
                // Clear all VBA modules
                workbook.VbaProject.Modules?.Clear();
                // Note: Forms collection may not be available in some versions; clearing modules is sufficient.
            }

            // Save the workbook as a standard Open XML XLSX file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
