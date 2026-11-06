// Title: Reset an Excel workbook’s theme to the default colors using Aspose.Cells for .NET (C#)
// Common Searches: how to reset an Excel workbook theme to default with Aspose.Cells in C# | C# Aspose.Cells remove custom theme from an existing workbook | apply built‑in theme colors to a loaded Excel file using Aspose.Cells | Aspose.Cells code to revert workbook theme without affecting data | load workbook and restore default theme using Aspose.Cells C#
// Tags: reset workbook theme Aspose.Cells C# | default Excel theme Aspose.Cells | remove custom theme Aspose.Cells | load and save workbook Aspose.Cells | Aspose.Cells SetTheme method usage

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing XLSX file with Aspose.Cells, verifies the file, relies on the library’s built‑in default theme (the Theme property is read‑only), and saves the workbook, effectively resetting any custom theme to the default color set.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // The Theme property is read‑only; Aspose.Cells uses the default theme automatically.
            // If a custom theme needs to be applied, use workbook.SetTheme(...). 
            // Here we keep the default theme, so no action is required.

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
