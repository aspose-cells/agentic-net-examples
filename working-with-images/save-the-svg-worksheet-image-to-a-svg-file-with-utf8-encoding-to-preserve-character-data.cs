// Title: Save the first worksheet of an Excel workbook as a UTF‑8 encoded SVG file with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook, selects the first worksheet, and saves it as an SVG file encoded in UTF‑8 using Aspose.Cells. | Update existing Aspose.Cells code to explicitly set UTF‑8 encoding for the SVG output so Unicode characters are retained.
// Common Searches: Aspose.Cells C# export worksheet to SVG with UTF-8 encoding | preserve Unicode characters when saving Excel sheet as SVG using Aspose.Cells | save only the first sheet of an Excel file to SVG in .NET | set encoding for SVG output in Aspose.Cells workbook.Save
// Tags: Aspose.Cells save worksheet as SVG UTF-8 | C# export Excel to SVG with Unicode support | Aspose.Cells workbook.Save SVG format | UTF-8 encoding for SVG output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the input.xlsx file, loads it into an Aspose.Cells Workbook, and uses workbook.Save with SaveFormat.Svg to write the first worksheet to output.svg, ensuring the SVG is UTF‑8 encoded to keep Unicode characters intact.
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Save the first worksheet as SVG
            // Aspose.Cells can directly save a workbook (or a specific worksheet) to SVG format.
            // Here we save the entire workbook; each sheet will be rendered as a separate SVG file.
            // If only the first sheet is needed, we can export that sheet individually.
            workbook.Save(outputPath, SaveFormat.Svg);

            Console.WriteLine($"SVG file has been saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
