// Title: C# example: read raw text from all merged cells with Aspose.Cells, concatenate them, and write the combined string to a summary cell
// Common Searches: Aspose.Cells C# read merged cell values and combine into one cell | How to get string value from merged cells using Aspose.Cells for .NET | Concatenate text from merged ranges in Excel with Aspose.Cells and save to another cell | Iterate merged cell areas in an Aspose.Cells workbook using C# | Write merged cells summary to a specific cell (e.g., B1) with Aspose.Cells .NET
// Tags: Aspose.Cells merged cells extraction C# | concatenate merged cell values Aspose.Cells | write summary cell Aspose.Cells workbook | iterate merged cell areas .NET | raw string value Aspose.Cells cell

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The code loads an existing Excel workbook, loops through every merged cell area on the first worksheet, retrieves the raw string from each area's top‑left cell, concatenates these strings with spaces, writes the resulting text to cell B1 as a summary, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Accumulate raw string values from merged cells
            StringBuilder concatenated = new StringBuilder();

            // Iterate through all merged cell areas in the worksheet
            foreach (CellArea area in sheet.Cells.MergedCells)
            {
                // The top‑left cell of the merged area holds the actual value
                Cell firstCell = sheet.Cells[area.StartRow, area.StartColumn];

                // Retrieve the raw string value (null‑safe)
                string value = firstCell.StringValue ?? string.Empty;

                // Append the value with a space delimiter
                if (concatenated.Length > 0)
                    concatenated.Append(' ');
                concatenated.Append(value);
            }

            // Write the concatenated result to a summary cell (e.g., B1)
            Cell summaryCell = sheet.Cells["B1"];
            summaryCell.PutValue(concatenated.ToString());

            // Save the workbook to a new file
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
