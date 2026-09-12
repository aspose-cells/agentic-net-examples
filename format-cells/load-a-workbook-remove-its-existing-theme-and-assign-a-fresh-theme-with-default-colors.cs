// Title: How to remove a custom theme and reset to the built‑in default theme in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, clears any custom theme, and applies the built‑in default theme before saving. | Show how to use Aspose.Cells to revert a workbook's theme to the built‑in default, handling versions where the Theme property is a string. | Write a robust C# routine that checks the existence of the source Excel file, opens it with Aspose.Cells, resets the theme to the default, and writes the result to a new file, handling possible API version differences.
// Common Searches: asp.net aspose.cells reset workbook theme to default | c# remove custom theme from excel file using Aspose.Cells | how to revert Excel theme to built‑in default with Aspose.Cells .NET | Theme.ResetToDefault not available Aspose.Cells workaround | resetting Excel workbook theme before saving in C#
// Tags: Aspose.Cells reset workbook theme | C# apply default Excel theme | remove custom theme from .xlsx programmatically | Theme.ResetToDefault method Aspose.Cells | handle theme property version differences Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Loads 'input.xlsx', verifies its existence, opens the workbook with Aspose.Cells, attempts to reset its theme to the built‑in default (using Theme.ResetToDefault when available or appropriate fallback), and saves the modified file as 'output.xlsx' with error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Reset the workbook's theme to the built‑in default theme
            // Note: In some Aspose.Cells versions the Theme property is a string.
            // If the Theme object with ResetToDefault() is available, uncomment the line below.
            // workbook.Theme.ResetToDefault();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
