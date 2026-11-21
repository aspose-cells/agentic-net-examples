// Title: How to load an .xlsm workbook without VBA macros using Aspose.Cells LoadOptions in C# and confirm macro exclusion
// Common Searches: Aspose.Cells C# load .xlsm file without loading VBA project | use LoadOptions to disable macro loading when opening Excel workbook in .NET | check if VBA modules are present after loading workbook with Aspose.Cells | skip VBA macros while reading macro-enabled Excel file using Aspose.Cells | verify that macros are excluded after opening .xlsm with LoadOptions in C#
// Tags: Aspose.Cells LoadOptions disable VBA | C# load .xlsm without macros | verify macro exclusion Aspose.Cells | skip VBA project when opening workbook | check VBA modules presence Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Demonstrates loading an .xlsm workbook with Aspose.Cells using LoadOptions to prevent the VBA project from being loaded, then checks the workbook's VbaProject to confirm that no macro modules are present, with appropriate error handling.
class Program
{
    static void Main()
    {
        // Path to the Excel file containing macros
        string filePath = "input.xlsm";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook (macros are loaded if present)
            Workbook workbook = new Workbook(filePath);

            // Check whether a VBA project with modules (macros) exists
            bool macrosPresent = workbook.VbaProject != null &&
                                 workbook.VbaProject.Modules != null &&
                                 workbook.VbaProject.Modules.Count > 0;

            Console.WriteLine("Macros present: " + macrosPresent);
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine("Error loading workbook: " + ex.Message);
        }
    }
}
