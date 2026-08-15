// Title: Copy Formatting Between Ranges in a Password‑Protected Excel Workbook Using Aspose.Cells for .NET (C#)
// Description: Loads a password‑protected workbook (protected.xlsx) with Aspose.Cells, creates the O1:O5 and P1:P5 ranges on the first worksheet, copies only the cell style via CopyStyle, and saves the result as output.xlsx.
// Keywords: Aspose.Cells | C# | CopyStyle | password protected workbook | range formatting | Excel automation | load workbook with password | copy cell style
// Common Searches: Aspose.Cells copy range formatting | load password protected Excel file C# Aspose.Cells | CopyStyle method example | copy only styles between cells Aspose.Cells | copy formatting without values Aspose.Cells
// Developer Intent: Load a password‑protected Excel file and copy only the formatting from cells O1:O5 to P1:P5 using Aspose.Cells for .NET.
// Use Cases: Apply a corporate template style to a new column in a secured workbook. | Transfer conditional formatting from a protected source sheet to another area within the same file. | Generate reports that reuse formatting from a locked master workbook while preserving original data.
// AI Prompts: Generate C# code that loads a password‑protected workbook and copies only the formatting from one range to another using Aspose.Cells. | Provide robust error‑handling examples for missing files and incorrect passwords when opening a workbook with Aspose.Cells. | Show how to copy formatting for multiple non‑contiguous ranges in a protected workbook using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads a password‑protected workbook (protected.xlsx) with Aspose.Cells, creates the O1:O5 and P1:P5 ranges on the first worksheet, copies only the cell style via CopyStyle, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "protected.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the password‑protected workbook
            var loadOptions = new LoadOptions
            {
                Password = "myPassword"
            };
            var workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Define source and destination ranges
            Aspose.Cells.Range sourceRange = sheet.Cells.CreateRange("O1:O5");
            Aspose.Cells.Range destinationRange = sheet.Cells.CreateRange("P1:P5");

            // Copy only the formatting from source to destination
            destinationRange.CopyStyle(sourceRange);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
