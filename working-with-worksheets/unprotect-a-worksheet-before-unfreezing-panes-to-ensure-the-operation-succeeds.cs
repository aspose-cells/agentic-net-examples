// Title: Unprotect an Excel worksheet and clear frozen panes using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to remove frozen panes from a protected worksheet | unprotect worksheet before calling FreezePanes with Aspose.Cells .NET | Aspose.Cells example to clear freeze panes after worksheet unprotect | remove worksheet protection and reset pane freezing in C# | Aspose.Cells unfreeze panes on a password‑less protected sheet
// Tags: worksheet unprotect Aspose.Cells C# | clear frozen panes Aspose.Cells .NET | reset FreezePanes on protected Excel sheet | Aspose.Cells unprotect then unfreeze worksheet | remove worksheet protection before FreezePanes

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel file, unprotects the first worksheet, clears any frozen panes by calling FreezePanes(0,0,0,0), and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Unprotect the worksheet (no password required)
            worksheet.Unprotect();

            // Clear any frozen panes (use four‑parameter overload with zeros)
            worksheet.FreezePanes(0, 0, 0, 0);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
