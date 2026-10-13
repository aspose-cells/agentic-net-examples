// Title: Count non‑empty cells in a given worksheet range with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, creates a range from a user‑provided address, and returns the count of cells whose Value is not null or whitespace. | Modify the example to treat numeric zero and boolean false as valid non‑empty entries when counting populated cells. | Write a reusable function that accepts a Workbook object and a range address, iterates over the cells using Aspose.Cells, and returns the total number of non‑empty cells while ignoring empty strings and nulls.
// Common Searches: Aspose.Cells C# count cells with data in specific range A1:C10 | How to get number of populated cells in an Excel worksheet using Aspose.Cells .NET | C# Aspose.Cells iterate over range and skip empty or whitespace cells | Calculate non‑blank cell count for a dynamic range address with Aspose.Cells
// Tags: count non‑empty cells Aspose.Cells C# | worksheet range cell enumeration Aspose.Cells | skip whitespace strings Aspose.Cells | populate cell count Excel .NET | range address processing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, creates a specified range (e.g., A1:C10), iterates through its cells with Aspose.Cells, and counts cells that contain a non‑null value and are not whitespace‑only strings.
class Program
{
    static void Main()
    {
        // Path to the input workbook
        string inputPath = "input.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
            return;
        }

        Workbook workbook;
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

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the range to evaluate (e.g., cells A1 to C10)
        string rangeAddress = "A1:C10";

        // Create the range (use fully qualified Aspose.Cells.Range to avoid ambiguity)
        Aspose.Cells.Range range = worksheet.Cells.CreateRange(rangeAddress);

        // Count non‑empty cells within the specified range
        int nonEmptyCount = 0;
        foreach (Cell cell in range)
        {
            // A cell is considered non‑empty if its Value is not null and, for strings, not whitespace
            if (cell.Value != null)
            {
                if (cell.Value is string str)
                {
                    if (!string.IsNullOrWhiteSpace(str))
                        nonEmptyCount++;
                }
                else
                {
                    nonEmptyCount++;
                }
            }
        }

        // Output the result
        Console.WriteLine($"Total non‑empty cells in range {rangeAddress}: {nonEmptyCount}");
    }
}
