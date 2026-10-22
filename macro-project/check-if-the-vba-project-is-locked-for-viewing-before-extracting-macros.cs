// Title: Determine if an Excel VBA project is locked before exporting macros with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells check VBA project lock status before calling ExportMacros | how to read the IsLocked property of a VBA project using Aspose.Cells .NET | export macros only when the VBA project is unlocked in an Excel workbook with C# | handle missing IsLocked property across different Aspose.Cells versions | detect locked VBA project in .xlsx files using Aspose.Cells API
// Tags: Aspose.Cells VBA lock detection C# | conditional ExportMacros execution | dynamic IsLocked property handling | VBA project lock status check .NET | macro extraction from unlocked workbook

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, verifies the presence of a VBA project, attempts to read its IsLocked property via dynamic invocation to support multiple API versions, and if the project is not locked, dynamically calls ExportMacros to save the macros to a .vba file, handling absent properties and export errors gracefully.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "extracted_macros.vba";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Check for a VBA project
            var vbaProject = workbook.VbaProject;
            if (vbaProject != null)
            {
                bool isLocked = false;

                // Attempt to read the IsLocked property via dynamic (covers different API versions)
                try
                {
                    isLocked = (bool)((dynamic)vbaProject).IsLocked;
                }
                catch
                {
                    // Property not available in this version; assume not locked
                }

                Console.WriteLine($"VBA project locked: {isLocked}");

                if (!isLocked)
                {
                    // Attempt to export macros via dynamic invocation
                    try
                    {
                        ((dynamic)vbaProject).ExportMacros(outputPath);
                        Console.WriteLine("Macros have been exported successfully.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to export macros: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine("The VBA project is locked. Macros cannot be extracted.");
                }
            }
            else
            {
                Console.WriteLine("No VBA project found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
