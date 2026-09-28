// Title: Update an Excel workbook’s theme to use a monospaced font (Consolas) with Aspose.Cells for .NET
// AI Prompts: Load an existing workbook and assign 'Consolas' as the default font using Aspose.Cells in C#. | Programmatically apply a code‑friendly typeface to the entire Excel theme and save the workbook with Aspose.Cells. | Change the workbook’s default font to a programming‑oriented font across all sheets via the Aspose.Cells API.
// Common Searches: Aspose.Cells C# change workbook font to Consolas | set programming font for all cells in an existing Excel file using Aspose.Cells | programmatically update Excel theme to a monospaced typeface with .NET | C# Aspose.Cells change default style font for new worksheets
// Tags: default workbook font Aspose.Cells C# | apply programming font to Excel theme programmatically | set workbook font to Consolas | modify Excel theme font scheme using Aspose.Cells | override workbook default style for all worksheets .NET

using Aspose.Cells;
using System;
using System.IO;

// Loads an existing Excel file, sets the workbook’s default font to the monospaced typeface Consolas, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Set the default font for the workbook to a monospaced font (e.g., Consolas)
            workbook.DefaultStyle.Font.Name = "Consolas";

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
