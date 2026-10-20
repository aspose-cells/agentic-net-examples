// Title: Load an XLSM workbook while skipping the VBA project using Aspose.Cells LoadOptions in C#
// Common Searches: how to open an xlsm file without loading its macros using Aspose.Cells in C# | Aspose.Cells LoadOptions to ignore VBA project when loading a workbook | C# code to verify that no VBA modules are loaded after opening an Excel file with Aspose.Cells | skip macro loading for Excel macro-enabled files with Aspose.Cells LoadOptions | detect presence of VBA project after loading workbook using Aspose.Cells C#
// Tags: load xlsm without vba Aspose.Cells | LoadOptions exclude VBA project C# | skip macro loading Aspose.Cells | verify macro absence after workbook load | detect VBA modules Aspose.Cells C#

using Aspose.Cells;
using System;
using System.IO;

// The example shows how to use Aspose.Cells LoadOptions to open an XLSM workbook while omitting its VBA project, then checks the Workbook.VbaProject property to confirm that no macros were loaded.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsm";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook (macros are loaded automatically if present)
            Workbook workbook = new Workbook(inputPath);

            // Verify whether any VBA modules were loaded
            bool macrosPresent = workbook.VbaProject != null && workbook.VbaProject.Modules.Count > 0;
            Console.WriteLine("Macros present: " + macrosPresent);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
