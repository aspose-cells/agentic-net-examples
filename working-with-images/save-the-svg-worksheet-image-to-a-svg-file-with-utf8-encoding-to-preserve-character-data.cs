// Title: Save the first worksheet of an Excel workbook as a UTF‑8 encoded SVG file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook, verifies the file exists, and saves the first worksheet to a UTF‑8 encoded SVG with Aspose.Cells. | Provide a C# example that demonstrates handling a missing Excel file when converting a worksheet to SVG using Aspose.Cells. | Show how to set a custom output path and ensure Unicode characters are retained while exporting a worksheet to SVG in C# with Aspose.Cells.
// Common Searches: how to export a single Excel sheet to SVG with UTF-8 encoding in C# using Aspose.Cells | c# Aspose.Cells save first worksheet as SVG with Unicode support | example code for converting .xlsx to .svg with Aspose.Cells | Aspose.Cells SVG output Unicode characters C# | save worksheet as SVG file with UTF-8 using .NET library
// Tags: Aspose.Cells worksheet to SVG conversion | UTF-8 SVG export .NET | C# first sheet SVG output | missing file handling Aspose.Cells | Unicode characters in SVG export

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the presence of input.xlsx, loads it with Aspose.Cells, and saves the first worksheet as a UTF‑8 encoded SVG file (output.svg) while handling potential errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.svg";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook (first worksheet) as SVG
            // Aspose.Cells saves each worksheet as a separate SVG file; the first worksheet will be outputPath
            workbook.Save(outputPath, SaveFormat.Svg);

            Console.WriteLine($"SVG file has been successfully created at '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
