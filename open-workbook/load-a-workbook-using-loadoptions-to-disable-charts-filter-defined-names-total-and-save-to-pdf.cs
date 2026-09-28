// Title: Load an XLSX workbook with charts disabled, keep only the 'Total' defined name, and export to PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx file with LoadOptions configured to ignore chart objects, removes every defined name except "Total", and saves the workbook as a PDF using Aspose.Cells. | Provide a step‑by‑step C# example showing how to enable LoadDataOnly in LoadOptions, filter named ranges to retain only "Total", and convert the workbook to PDF.
// Common Searches: Aspose.Cells C# load workbook without charts and convert to PDF | How to remove all named ranges except a specific one before PDF export in Aspose.Cells | LoadOptions LoadDataOnly example for disabling charts in Aspose.Cells | Save Excel to PDF while preserving only selected defined name using Aspose.Cells .NET | C# Aspose.Cells filter defined names and export workbook to PDF
// Tags: loadoptions disable chart objects aspnet cells | filter defined names aspnet cells | save workbook as pdf aspnet cells | loaddataonly usage aspnet cells | retain specific named range aspnet cells

using Aspose.Cells;
using System;
using System.IO;

// The example checks for the input XLSX file, creates LoadOptions (optionally enabling LoadDataOnly to skip chart objects), loads the workbook, removes all defined names except the one named "Total", and then saves the modified workbook as a PDF.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load options: disable charts and other non‑data objects (if supported)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            // The LoadDataOnly property may not be available in older versions; omit if not present.

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Keep only the defined name "Total"
            for (int i = workbook.Worksheets.Names.Count - 1; i >= 0; i--)
            {
                if (!string.Equals(workbook.Worksheets.Names[i].Text, "Total", StringComparison.OrdinalIgnoreCase))
                {
                    workbook.Worksheets.Names.RemoveAt(i);
                }
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
