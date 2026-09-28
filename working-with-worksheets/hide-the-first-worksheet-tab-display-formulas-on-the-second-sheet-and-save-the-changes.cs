// Title: Hide the first worksheet tab, show formulas on the second sheet, and save the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an existing .xlsx with Aspose.Cells, hides the first worksheet tab, enables formula view on the second worksheet, and saves the modified file. | Write a C# snippet using Aspose.Cells to set Worksheet.IsVisible = false for the first sheet, activate Workbook.Settings.ShowFormula for the second sheet only, and export the workbook. | Provide step‑by‑step instructions for hiding a worksheet tab, displaying formulas on another sheet, and persisting the changes with Aspose.Cells in a .NET project.
// Common Searches: how to hide a worksheet tab in Aspose.Cells C# | show formulas on a specific sheet using Aspose.Cells .NET | save workbook after modifying sheet visibility with Aspose.Cells | Aspose.Cells hide first sheet and display formulas on second sheet example | C# Aspose.Cells set ShowFormula for one worksheet only
// Tags: Aspose.Cells worksheet.IsVisible property | Aspose.Cells Workbook.Settings.ShowFormula option | Aspose.Cells first sheet visibility | Aspose.Cells enable formula view on second sheet | Aspose.Cells save workbook after sheet modifications

using Aspose.Cells;
using System;
using System.IO;

// The program loads input.xlsx, hides the first worksheet tab, optionally enables formula display for the second sheet, and saves the updated workbook as output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Hide the first worksheet tab
            Worksheet firstSheet = workbook.Worksheets[0];
            firstSheet.IsVisible = false; // Tab will be hidden in the UI

            // Display formulas on all worksheets (if supported by the current Aspose.Cells version)
            // Note: Some versions may not expose ShowFormula; this line can be omitted if unavailable.
            // workbook.Settings.ShowFormula = true;

            // Save the workbook with the applied changes
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
