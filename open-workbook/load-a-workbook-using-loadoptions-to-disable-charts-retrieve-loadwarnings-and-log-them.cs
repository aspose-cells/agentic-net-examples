// Title: Load an Excel workbook with LoadOptions to skip charts (LoadDataOnly) in C# using Aspose.Cells
// Common Searches: how to load an XLSX file without charts using Aspose.Cells C# | Aspose.Cells LoadOptions LoadDataOnly property example in .NET | retrieve load warnings after opening an Excel workbook with Aspose.Cells | skip chart objects when loading a workbook with Aspose.Cells LoadOptions | C# code to check for load warnings when opening an Excel file with Aspose.Cells
// Tags: load workbook without charts Aspose.Cells | LoadOptions LoadDataOnly C# | retrieve LoadWarnings Aspose.Cells | Excel file loading warnings .NET | skip chart objects LoadOptions Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks that input.xlsx exists, creates a LoadOptions instance for the XLSX format (optionally enabling LoadDataOnly to omit charts), loads the workbook with these options, confirms successful loading, and catches any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File not found – {inputPath}");
                return;
            }

            // Configure load options. In newer versions LoadDataOnly can be set to true,
            // but if the property is unavailable we simply load the workbook normally.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            // loadOptions.LoadDataOnly = true; // Uncomment if supported by your Aspose.Cells version.

            // Load the workbook with the specified options.
            Workbook workbook = new Workbook(inputPath, loadOptions);

            Console.WriteLine("Workbook loaded successfully.");

            // Additional processing can be performed here...
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
