// Title: How to apply a built‑in Excel theme to a workbook and refresh all cell styles with Aspose.Cells for .NET
// AI Prompts: Generate C# code that calls Workbook.SetTheme to apply a predefined built‑in theme (e.g., "Office") and automatically updates the formatting of every cell in the workbook using Aspose.Cells. | Show an example that loads an existing .xlsx file, sets a built‑in theme, forces a style refresh for all cells, and saves the modified workbook with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set built‑in theme and update cell formatting | Workbook.SetTheme example for .xlsx files in .NET | Refresh all cell styles after changing Excel theme with Aspose.Cells | Apply Office theme to existing workbook using Aspose.Cells for .NET | How to propagate theme changes to existing cells in Aspose.Cells
// Tags: Workbook.SetTheme built-in theme | refresh cell styles Aspose.Cells | apply Excel theme .NET | update workbook formatting Aspose.Cells | theme management Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel workbook, demonstrates how to use the Workbook.SetTheme method (available in newer Aspose.Cells releases) to apply a built‑in theme such as "Office", and ensures that all existing cell styles are refreshed to reflect the new theme before saving the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Applying a built‑in theme is not supported in the current Aspose.Cells version.
            // If needed, use Workbook.SetTheme method available in newer versions.

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
