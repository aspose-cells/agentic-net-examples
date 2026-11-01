// Title: Load an Excel workbook with automatic row height autofit using Aspose.Cells LoadOptions.AutoFitterOptions.OnlyAuto in C#
// Common Searches: C# Aspose.Cells load workbook with AutoFitterOptions to autofit row heights | How to enable automatic row height adjustment when opening an Excel file with Aspose.Cells | Using LoadOptions.AutoFitterOptions.OnlyAuto true to fit rows on workbook load in .NET | Aspose.Cells load existing .xlsx and auto‑fit all rows without calling AutoFitRows | Create a workbook that automatically adjusts row heights on load with Aspose.Cells
// Tags: Aspose.Cells LoadOptions AutoFitterOptions | automatic row height autofit on load | C# load Excel workbook with OnlyAuto | row height adjustment during workbook import | Excel file loading with auto‑fit rows

using Aspose.Cells;
using System;
using System.IO;

// The example demonstrates how to configure LoadOptions with AutoFitterOptions.OnlyAuto set to true so that Aspose.Cells automatically adjusts all row heights when an existing Excel workbook is loaded, then saves the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Check if the input file exists; if not, create a new workbook.
            if (File.Exists(inputPath))
            {
                // Configure load options to automatically adjust all row heights on load.
                LoadOptions loadOptions = new LoadOptions
                {
                    AutoFitterOptions = new AutoFitterOptions
                    {
                        OnlyAuto = true
                    }
                };

                // Load the workbook with the specified options.
                workbook = new Workbook(inputPath, loadOptions);
            }
            else
            {
                // Input file not found – create an empty workbook.
                workbook = new Workbook();
            }

            // Save the workbook (row heights are already adjusted if loaded).
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors (e.g., password protection, I/O issues).
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
