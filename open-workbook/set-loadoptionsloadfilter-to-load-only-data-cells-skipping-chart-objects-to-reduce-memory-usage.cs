// Title: Load only worksheet data cells while skipping charts with Aspose.Cells LoadFilter in C#
// Common Searches: C# load Excel workbook without charts using Aspose.Cells LoadFilter | How to reduce memory usage when opening large XLSX files with Aspose.Cells | Aspose.Cells LoadOptions LoadDataOnly example for C# | Skip chart objects while loading workbook in Aspose.Cells C# | Load only cell values from XLSX using Aspose.Cells in .NET
// Tags: Aspose.Cells LoadFilter data‑only option | load worksheet data cells only C# | reduce memory usage loading XLSX Aspose | exclude chart objects Aspose.Cells | LoadOptions LoadDataOnly API

using System;
using System.IO;
using Aspose.Cells;

// This C# example demonstrates how to configure LoadOptions with a LoadFilter set to LoadDataOnly so that Aspose.Cells loads only the worksheet data cells and skips chart objects, lowering memory consumption when opening an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Create load options for an XLSX file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // If the Aspose.Cells version supports LoadFilterOptions, you can uncomment the following line
            // loadOptions.LoadFilter = new LoadFilter(LoadFilterOptions.LoadDataOnly);

            // Load the workbook using the configured load options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // (Optional) Perform any data processing here

            // Save the workbook; charts are not loaded if LoadFilter was applied
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
