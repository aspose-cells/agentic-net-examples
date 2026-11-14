// Title: Preserve Excel formulas by setting calculation mode to Automatic before and after merging cells with Aspose.Cells for .NET (C#)
// Common Searches: how to keep formulas from breaking when merging cells using Aspose.Cells C# | set workbook calculation mode to automatic before merging cells Aspose.Cells .NET | Aspose.Cells preserve formula integrity during cell merge in C# | C# merge cells A1:C3 without losing formulas Aspose.Cells | automatic calculation mode after merging cells Aspose.Cells example
// Tags: automatic calculation mode Aspose.Cells | merge cells preserve formulas C# | Aspose.Cells workbook calculation settings | cell range merge Excel Aspose.Cells | C# Aspose.Cells formula integrity

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, sets the calculation mode to Automatic before merging the A1:C3 range, performs the merge, re‑applies Automatic mode to ensure formulas stay intact, and saves the workbook, handling any errors that may occur.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Merge cells A1:C3 in the first worksheet (rows 0‑2, columns 0‑2)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells.Merge(0, 0, 3, 3);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
