// Title: Read a formatted Excel cell, replace commas with semicolons, and write the result back using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells read displayed cell text and replace commas with semicolons | How to update an Excel cell value without losing its formatting using Aspose.Cells | Replace commas in an Excel cell string programmatically with Aspose.Cells .NET | Write modified string back to the same cell and save workbook in Aspose.Cells | Aspose.Cells get cell string value, perform text replace, and preserve formatting
// Tags: formatted cell text retrieval Aspose.Cells | cell string replacement .NET | preserve cell formatting during value update C# | write back modified cell value Aspose.Cells | save workbook after cell modification .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, reads the displayed text of cell A1, replaces every comma with a semicolon, writes the updated string back to the same cell while keeping its original formatting, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook (no password supplied). If the file is password‑protected,
            // an exception will be caught below.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Target cell (adjust address as required)
            Cell cell = worksheet.Cells["A1"];

            // Retrieve the cell's displayed text
            string formattedText = cell.StringValue;

            // Replace commas with semicolons
            string modifiedText = formattedText.Replace(",", ";");

            // Write the modified string back to the same cell
            cell.PutValue(modifiedText);

            // Save the workbook with the changes
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            // Handles Aspose.Cells specific errors (e.g., password protection)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
