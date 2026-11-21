// Title: Export an Excel worksheet to an SVG file with UTF‑8 encoding using Aspose.Cells for .NET
// Common Searches: how to save an Excel worksheet as an SVG file with UTF-8 encoding in C# using Aspose.Cells | Aspose.Cells C# export first worksheet to SVG preserving Unicode characters | convert .xlsx to .svg with UTF-8 output using Aspose.Cells .NET | C# code example for workbook.Save with SaveFormat.Svg and UTF-8 | export Excel sheet to SVG while keeping special characters using Aspose.Cells
// Tags: Aspose.Cells export worksheet to SVG | C# SaveFormat.Svg UTF-8 output | Excel to SVG conversion with Unicode preservation | Workbook.Save SVG example .NET | UTF-8 encoded SVG export from Excel | Aspose.Cells SVG file generation C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook, verifies the input file, and uses Workbook.Save with SaveFormat.Svg to export the first worksheet to a UTF‑8 encoded SVG file, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.svg";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Export the first worksheet to SVG
            // SaveFormat.Svg creates an SVG file for each worksheet; we use the first one.
            workbook.Save(outputPath, SaveFormat.Svg);

            Console.WriteLine($"Worksheet exported to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
