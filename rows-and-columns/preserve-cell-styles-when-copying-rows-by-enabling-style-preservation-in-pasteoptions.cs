// Title: Copy an Excel row to another location while preserving cell styles using Aspose.Cells for .NET
// AI Prompts: Generate C# code that copies row 2 to row 6 in an existing .xlsx workbook, ensuring all cell formatting (fonts, colors, borders) is retained using Aspose.Cells PasteOptions. | Show how to use Aspose.Cells Range.Copy with PasteOptions to duplicate a row while preserving its style in a .NET console application. | Write a C# snippet that loads an Excel file, copies a source row to a destination row, and saves the file, keeping the original cell styles intact.
// Common Searches: Aspose.Cells .NET copy row preserving formatting to another row index | How to keep cell styles when copying rows with Aspose.Cells in C# | Using PasteOptions to retain formatting while copying Excel rows in Aspose.Cells | Copy entire row with formatting using Aspose.Cells Range.Copy C# example
// Tags: Aspose.Cells row copy style retention | PasteOptions keep cell formatting .NET | range.Copy preserve cell styles Aspose.Cells | C# duplicate Excel row with formatting | Excel row copy preserving styles Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, defines source and destination row ranges, copies the source row to the target row while preserving all cell formatting using Aspose.Cells, and saves the result to a new file.
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
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Define source and destination row indices (0‑based)
            int sourceRow = 1;        // Row to copy (second row)
            int destinationRow = 5;   // Row where the copy will be placed

            // Determine the number of columns to copy (entire used range)
            int maxColumn = sheet.Cells.MaxColumn;

            // Create source range covering the entire source row
            Aspose.Cells.Range sourceRange = sheet.Cells.CreateRange(sourceRow, 0, 1, maxColumn + 1);

            // Create destination range with the same width
            Aspose.Cells.Range destRange = sheet.Cells.CreateRange(destinationRow, 0, 1, maxColumn + 1);

            // Copy the source range to the destination range preserving all data and styles
            sourceRange.Copy(destRange);

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
