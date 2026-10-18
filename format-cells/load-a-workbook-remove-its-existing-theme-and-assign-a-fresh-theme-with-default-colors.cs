// Title: Load an Excel workbook and save it unchanged while noting the lack of a direct theme‑reset API in Aspose.Cells for .NET
// Common Searches: how to reset an Excel workbook theme using Aspose.Cells in C# | Aspose.Cells SetTheme method example for applying a new theme | remove existing theme from an .xlsx file with Aspose.Cells .NET | load and save Excel workbook without modifying theme Aspose.Cells | Aspose.Cells limitation for resetting workbook theme to default
// Tags: Aspose.Cells workbook theme reset limitation | Aspose.Cells SetTheme API usage | C# load and save Excel workbook with Aspose.Cells | programmatic removal of Excel theme .NET | apply default theme to workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample checks for the presence of input.xlsx, loads it into an Aspose.Cells Workbook, explains that Aspose.Cells does not provide a direct API to reset a workbook's theme, and then saves the workbook unchanged to output.xlsx while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Aspose.Cells does not provide a direct API to reset to a default theme.
            // If needed, you can apply a specific theme file using SetTheme(string themePath).
            // Here we simply save the workbook without modifying the theme.

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
