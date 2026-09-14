// Title: Count cells with custom number formats in each worksheet of an Excel file using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx workbook with Aspose.Cells, iterates all worksheets, and counts cells whose Style.Custom property contains a custom number format. | Extend the sample to gather and print the distinct custom number format strings present in each worksheet. | Add robust try‑catch handling so that cells causing exceptions are skipped while still contributing to the custom‑format count.
// Common Searches: aspnet count cells with custom number format per worksheet using Aspose.Cells | how to detect custom number formats in Excel with Aspose.Cells C# | enumerate used range and find custom number format strings Aspose.Cells .NET | C# Aspose.Cells get count of cells that have a custom number format in each sheet
// Tags: custom number format detection Aspose.Cells | worksheet cell enumeration Aspose.Cells | Style.Custom property usage .NET | count custom formatted cells per sheet | error‑tolerant cell scanning Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, iterates each worksheet's used range, checks the Style.Custom property of every cell, counts cells with a non‑empty custom number format, and prints the count per worksheet while handling missing files and cell‑level exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            int customFormatCount = 0;

            // Determine the used range of the sheet
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Scan cells only if the sheet contains data
            if (maxRow >= 0 && maxCol >= 0)
            {
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        try
                        {
                            Cell cell = sheet.Cells[row, col];
                            Style style = cell.GetStyle();

                            // A custom number format is indicated by a non‑empty Custom string
                            string customFormat = style.Custom;
                            if (!string.IsNullOrEmpty(customFormat))
                            {
                                customFormatCount++;
                            }
                        }
                        catch (Exception cellEx)
                        {
                            // Log cell‑level errors but continue processing
                            Console.WriteLine($"Warning: Unable to process cell [{row}, {col}] in sheet \"{sheet.Name}\": {cellEx.Message}");
                        }
                    }
                }
            }

            // Report the count for the current worksheet
            Console.WriteLine($"Worksheet \"{sheet.Name}\": {customFormatCount} cells with custom number formats.");
        }
    }
}
