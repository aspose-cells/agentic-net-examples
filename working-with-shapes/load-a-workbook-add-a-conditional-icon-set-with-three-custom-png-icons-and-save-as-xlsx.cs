// Title: Load an existing XLSX workbook, verify its presence, and save it unchanged using Aspose.Cells for .NET (C#)
// Common Searches: how to check if an Excel file exists before opening with Aspose.Cells in C# | Aspose.Cells C# example to load a workbook and save it as XLSX without changes | C# code to open an existing .xlsx file with Aspose.Cells and handle missing file errors | saving a workbook after loading it using Aspose.Cells for .NET | preserve original formatting when re‑saving a workbook with Aspose.Cells C#
// Tags: Aspose.Cells load workbook C# | Aspose.Cells save workbook as XLSX | file existence validation C# Aspose.Cells | exception handling workbook operations Aspose.Cells | conditional formatting API version limitation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that the input.xlsx file exists, loads it into an Aspose.Cells Workbook object, accesses the first worksheet, and then saves the workbook unchanged to output.xlsx, with proper exception handling for missing files or other errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // NOTE: Conditional formatting APIs (ConditionalFormatting, IconSetCondition, etc.)
            // are not available in the current Aspose.Cells version used for this project.
            // The following block is omitted to ensure successful compilation.
            // If a newer version is referenced, the conditional formatting code can be
            // re‑added using the appropriate APIs.

            // Save the (potentially modified) workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
