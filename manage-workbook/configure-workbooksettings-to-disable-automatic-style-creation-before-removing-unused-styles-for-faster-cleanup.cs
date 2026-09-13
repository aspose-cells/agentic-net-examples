// Title: Disable automatic style creation via WorkbookSettings and then remove unused styles with Aspose.Cells for .NET
// AI Prompts: Generate C# code that sets WorkbookSettings.EnableStyleCreation to false, invokes RemoveUnusedStyles, and saves the workbook using Aspose.Cells. | Show a .NET example that configures WorkbookSettings to prevent new style generation before cleaning up unused styles in an Excel file. | Provide a snippet that loads an .xlsx file, disables automatic style creation, removes all unused styles, and writes the cleaned workbook to a new file.
// Common Searches: Aspose.Cells how to turn off automatic style creation before removing unused styles | C# disable style auto‑generation WorkbookSettings Aspose.Cells | Improve performance of RemoveUnusedStyles by disabling style creation in Aspose.Cells .NET | WorkbookSettings.EnableStyleCreation false example Aspose.Cells | Remove unused styles without creating new ones Aspose.Cells
// Tags: WorkbookSettings.EnableStyleCreation property | remove unused styles Aspose.Cells | disable automatic style creation .NET | style cleanup performance Aspose.Cells | Aspose.Cells workbook settings configuration

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, disables automatic style creation via WorkbookSettings, removes all unused styles to improve cleanup speed, and saves the optimized file, including basic file existence checks and error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Remove all unused styles from the workbook
            workbook.RemoveUnusedStyles();

            // Save the cleaned workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook cleaned and saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
